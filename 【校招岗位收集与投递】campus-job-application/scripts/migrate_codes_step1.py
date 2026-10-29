#!/usr/bin/env python3
"""一次性迁移：第 1 步固定编码（公司 C / 站点 S / 岗位 J）。验证通过后与回滚点一起删除。

默认只读，输出迁移报告（--report 写 JSON）；用户确认报告后 --apply --quote "<用户原话>" 才写入。

迁移内容：
  主表        给公司名、岗位行分配编码；按“公司 + 域名”建站点（含只在公司文件夹/登录配置里出现的站点）；
              公司.json 里的 投递链接 / 顺序 / 城市选项 搬进主表岗位行（投递链接、志愿顺序进确认台账）
  确认台账    状态确认、链接例外的行身份 → 编码（按 identity_changes 换过链接的先接上）；删去 identity_changes
  交接区      生成 岗位交接/公司/<C>.json；当前批次搬进各家「批次」，状态回到“等待正向评审”（内容指纹变了，须用户重新确认一次）；
              删除旧 公司.json 与 投递批次.json（登录配置在第 2 步前原地保留）
  Excel       受控首写：Excel 加上编码列（状态逐行一致才允许）
  投递组件私有 站点资产键绑定 templates/site-asset-keys.json；模板绑定 岗位ID → 岗位编码；
              作废运行队列与公司批次的账号占用（任务只能按新编码重新导入）
"""
import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tarfile
import time
from pathlib import Path
from urllib.parse import urlparse

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import json_excel_sync as S  # noqa: E402
import master_codes as M  # noqa: E402
import company_store as C  # noqa: E402

LEGACY_ID = re.compile(r'^[a-z0-9][a-z0-9_-]{0,79}$')
ASSET = re.compile(r'^(登录|表单定义)(?:@([^/\\]+))?\.json$')


def slug(s):
    return re.sub(r'^-+|-+$', '', re.sub(r'[^a-z0-9]+', '-', s.lower()))


def origin_of(url):
    p = urlparse(str(url or ''))
    return f'{p.scheme}://{p.netloc}' if p.scheme and p.netloc else None


def old_row_key(sheet, row, idx):
    """迁移前的行身份（地区表 + 公司名 + 岗位链接 + 岗位名），只在本脚本里用来把旧引用对上行，之后不再使用。"""
    company = S.cell_value(row, idx[S.COMPANY_COL])
    link = S.cell_value(row, idx[S.LINK_COL]) if S.LINK_COL in idx else None
    name = S.cell_value(row, idx[S.NAME_COL]) if S.NAME_COL in idx else None
    k = {'sheet': sheet, 'company': company}
    if link:
        k['link'] = link
    if name:
        k['title'] = name
    if not link and not name:
        k['placeholder'] = True
    return k


def old_key(d):
    k = {'sheet': d['sheet'], 'company': d['company']}
    if d.get('link'):
        k['link'] = d['link']
    if d.get('title'):
        k['title'] = d['title']
    if not d.get('link') and not d.get('title'):
        k['placeholder'] = True
    return k


class Migration:
    def __init__(self, ws, labels, private_dir, queue_dir):
        self.ws = Path(ws).resolve()
        self.ctx = S.Ctx(self.ws, labels)
        self.store = C.Store(self.ws, labels)
        self.private = self.ws / private_dir
        self.queue_dir = self.ws / queue_dir
        self.handoff = self.store.handoff
        self.base = self.store.base
        self.report = {'errors': [], 'warnings': []}

    def err(self, msg):
        self.report['errors'].append(msg)

    def warn(self, msg):
        self.report['warnings'].append(msg)

    # ---------- 读 ----------
    def load(self):
        self.master = self.ctx.load_master()
        self.ledger = self.ctx.ledger()
        self.ctx.ledger = lambda: self.ledger  # 校验读内存中的新台账
        if M.COUNTERS in self.master:
            raise SystemExit('主表已有编码计数器：编码迁移已做过，不能重复执行')
        ew = self.master.get('excel_written') or {}
        self.report['excelInSyncBeforeMigration'] = ew.get('data_sha256') == self.master.get('data_sha256')
        if not self.report['excelInSyncBeforeMigration']:
            self.err('主表有未写入 Excel 的改动：先按正常流程写入（或撤销），Excel 与主表一致后才能迁移')
        self.index = {}
        for name, s, idx, row in S.iter_status_rows(self.ctx, self.master):
            k = S.key_id(old_row_key(name, row, idx))
            if k in self.index:
                self.err(f'主表有两行旧身份相同：{k}')
            self.index[k] = (name, row, s)
        # 旧公司文件夹
        self.firms = []
        for f in sorted(self.base.glob('*/*/公司.json')):
            doc = json.loads(f.read_text(encoding='utf-8'))
            logins = {}
            primary = None
            for a in sorted(f.parent.iterdir()):
                m = ASSET.match(a.name)
                if not m or m.group(1) != '登录':
                    continue
                asset = json.loads(a.read_text(encoding='utf-8'))
                o = origin_of(asset.get('origin')) or ('https://' + m.group(2) if m.group(2) else None)
                if not m.group(2):
                    primary = o
                logins[o] = {'file': a, 'id': asset.get('companyId')}
            self.firms.append({'file': f, 'doc': doc, 'logins': logins, 'primary': primary})
        bf = self.handoff / '投递批次.json'
        self.batch = json.loads(bf.read_text(encoding='utf-8')) if bf.exists() else None
        of = self.handoff / '投递回传.json'
        self.outcomes = json.loads(of.read_text(encoding='utf-8')) if of.exists() else None
        tpl = self.private / 'templates'
        self.role_bindings = json.loads((tpl / 'role-bindings.json').read_text(encoding='utf-8')) if (tpl / 'role-bindings.json').exists() else {}
        self.company_bindings = json.loads((tpl / 'company-bindings.json').read_text(encoding='utf-8')) if (tpl / 'company-bindings.json').exists() else {}
        if (tpl / 'site-asset-keys.json').exists():
            raise SystemExit('已有站点资产键绑定：编码迁移已做过')
        ow = self.private / 'runs' / 'company-channel-owners.json'
        self.owners = json.loads(ow.read_text(encoding='utf-8')) if ow.exists() else {}
        self.queue_files = sorted(self.queue_dir.glob('*.json')) if self.queue_dir.exists() else []

    # ---------- 旧身份链 ----------
    def resolve(self, key):
        """旧行身份 → 当前行（先沿 identity_changes 走到最新身份）。"""
        k = S.key_id(old_key(key))
        seen = set()
        while k in self.chain and k not in seen:
            seen.add(k)
            k = self.chain[k]
        hit = self.index.get(k)
        return (hit[0], hit[1], S.headers(hit[2])) if hit else None  # 表头在分配编码后才有编码列，现取

    def code_of(self, hit):
        name, row, idx = hit
        col = M.ROLE_CODE_COL if S.row_kind(row, idx) == 'role' else M.COMPANY_CODE_COL
        return S.cell_value(row, idx[col])

    # ---------- 编码 ----------
    def assign(self):
        self.chain = {S.key_id(old_key(e['old_key'])): S.key_id(old_key(e['new_key'])) for e in self.ledger.get('identity_changes', [])}
        self.report['assigned'] = M.assign_missing_codes(self.master, self.ctx.sheets)
        # 旧目录项编号（站点资产键）：主站点用公司.json 的 companyId，其他站点用其登录配置的 companyId，否则 <companyId>--<域名>
        self.site_keys, self.legacy_company, sites_added = {}, {}, []
        for firm in self.firms:
            doc = firm['doc']
            ccode = M.resolve_company(self.master, doc['公司'])
            if not ccode:
                self.err(f"公司文件夹 {firm['file'].parent.name} 的公司名「{doc['公司']}」在主表公司主数据里对不上")
                continue
            firm['code'] = ccode
            primary = firm['primary']
            origins = list(dict.fromkeys([o for o in [primary, *(doc.get('origins') or []), *firm['logins']] if o]))
            primary = primary or (origins[0] if origins else None)
            for o in origins:
                if o == primary:
                    legacy = doc['companyId']
                else:
                    lid = (firm['logins'].get(o) or {}).get('id')
                    legacy = lid if lid and LEGACY_ID.match(lid) and lid != doc['companyId'] else (doc['companyId'] + '--' + slug(urlparse(o).netloc))[:80]
                before = M.site_of(self.master, ccode, o)
                scode = M.ensure_site(self.master, ccode, o)
                if not before:
                    sites_added.append(f"{doc['公司']}（{ccode}）{o} → {scode}")
                if scode in self.site_keys and self.site_keys[scode] != legacy:
                    self.err(f'站点 {scode} 对上两个旧目录项：{self.site_keys[scode]}、{legacy}')
                self.site_keys[scode] = legacy
                self.legacy_company[legacy] = doc['companyId']
            # 登录配置里写的 companyId 只作核对
            for o, lg in firm['logins'].items():
                if lg['id'] and lg['id'] not in self.site_keys.values():
                    self.warn(f"{firm['file'].parent.name} 的 {lg['file'].name} 的 companyId「{lg['id']}」不是任何站点的资产键")
        self.report['sitesFromCompanyFolders'] = sites_added
        self.report['siteAssetKeys'] = dict(sorted(self.site_keys.items()))

    # ---------- 公司.json 里的岗位值 ----------
    def move_role_values(self, quote, at):
        conf = self.ledger.setdefault('value_confirmations', [])
        self.role_ids, moved, unmatched = {}, [], []
        for firm in self.firms:
            doc = firm['doc']
            oq = doc.get('志愿顺序确认') or {}
            for r in doc.get('roles') or []:
                hit = self.resolve(r.get('主表行身份') or {})
                if not hit:
                    unmatched.append(f"{doc['公司']}｜{(r.get('主表行身份') or {}).get('title')}（岗位ID {r.get('岗位ID')}）")
                    continue
                name, row, idx = hit
                j = self.code_of(hit)
                self.role_ids[(doc['companyId'], str(r.get('岗位ID')))] = j
                label = f"{name}｜{doc['公司']}｜{S.cell_value(row, idx[S.NAME_COL])}（{j}）"
                if r.get('投递链接'):
                    site = S.cell_value(row, idx[M.SITE_CODE_COL])
                    if origin_of(r['投递链接']) != M.sites(self.master).get(site, {}).get('站点'):
                        self.err(f'{label} 的投递链接与岗位站点不同域：{r["投递链接"]}')
                    self.set_value(row, idx, j, M.APPLY_COL, r['投递链接'], label, quote, at, conf)
                    moved.append(f"{label}：投递链接「空」→「{r['投递链接']}」")
                if r.get('顺序') is not None:
                    q = oq.get('userQuote') or quote
                    self.set_value(row, idx, j, M.ORDER_COL, int(r['顺序']), label, q, oq.get('confirmedAt') or at, conf,
                                   ref='迁移自 公司.json 志愿顺序确认' if oq else '编码迁移')
                    moved.append(f"{label}：志愿顺序「空」→「{r['顺序']}」")
                if r.get('城市选项'):
                    v = M.ALIAS_SEP.join(str(x) for x in r['城市选项'])
                    S.set_cell(row, idx[M.CITY_COL], v)
                    moved.append(f"{label}：城市选项「空」→「{v}」")
        self.report['roleValuesMoved'] = moved
        if unmatched:
            self.warn('公司.json 里对不上主表行的岗位（这些岗位的投递链接/顺序/城市选项不迁移）：' + '；'.join(unmatched))
        self.report['unmatchedCompanyFileRoles'] = unmatched

    def set_value(self, row, idx, code, field, value, label, quote, at, conf, ref='编码迁移'):
        col = idx[field]
        S.set_cell(row, col, value)
        if field == M.APPLY_COL:
            next(c for c in row['cells'] if c['column'] == col)['hyperlink'] = {'target': value, 'location': None}
        conf[:] = [e for e in conf if (e['code'], e['field']) != (code, field)]
        conf.append({'code': code, 'field': field, 'value': value, 'shown_to_user': f'{label}：{field}「空」→「{value}」',
                     'user_quote': quote, 'confirmation_ref': ref, 'confirmed_at': at})

    # ---------- 台账 ----------
    def convert_ledger(self):
        out = {}
        for part in ('entries', 'link_exceptions'):
            kept, dropped, missing = [], [], []
            for e in self.ledger.get(part, []):
                hit = self.resolve(e['key'])
                if hit:
                    e = dict(e, key={'sheet': hit[0], 'code': self.code_of(hit)})
                    kept.append(e)
                elif part == 'entries' and (e.get('applied_at') or e.get('superseded_at')):
                    dropped.append(e.get('shown_to_user') or S.key_id(e['key']))
                else:
                    missing.append(e.get('shown_to_user') or S.key_id(e['key']))
            self.ledger[part] = kept
            out[part] = {'converted': len(kept), 'droppedFinishedHistory': len(dropped), 'unresolved': missing}
            if dropped:
                self.warn(f'{part}：{len(dropped)} 条已写入/已作废的旧确认对不上任何现有行（行已删除），不再保留')
            for m in missing:
                self.err(f'{part} 有未完成的记录对不上现有行：{m}')
        out['identityChangesRemoved'] = len(self.ledger.pop('identity_changes', []))
        self.report['ledger'] = out

    # ---------- 批次 ----------
    def convert_batch(self):
        if not self.batch:
            self.report['batch'] = None
            return
        b = self.batch
        by_legacy = {f['doc']['companyId']: f for f in self.firms}
        plan, roles_of = [], {}
        for i, c in enumerate(b['companies'], start=1):
            firm = by_legacy.get(c['companyId'])
            if not firm or 'code' not in firm:
                self.err(f"批次里的公司 {c.get('公司')}（{c['companyId']}）对不上公司文件夹")
                continue
            js = []
            for r in c['岗位']:
                j = self.role_ids.get((c['companyId'], str(r['岗位ID'])))
                if not j:
                    self.err(f"批次岗位 {c.get('公司')}｜{r.get('岗位名称')}（{r['岗位ID']}）对不上主表行")
                    continue
                js.append(j)
                roles_of[f"{c['companyId']}|{r['岗位ID']}"] = (firm['code'], j)
                if r.get('投递链接'):
                    hit = self.find_role(j)
                    cur = S.cell_value(hit[1], hit[2][M.APPLY_COL]) if hit else None
                    if cur and cur != r['投递链接']:
                        self.err(f"批次里 {j} 的投递链接与公司.json 不同：{r['投递链接']} / {cur}")
                    elif not cur and hit:
                        self.warn(f"批次里 {j} 有投递链接而公司.json 没有：{r['投递链接']}（不迁移，等投递端重新观察）")
            plan.append({'code': firm['code'], 'name': c.get('公司'), 'order': i, 'roles': js})
        prog = b.get('progress') or {}
        done = {roles_of[k][1] for k in prog.get('processedOutcomeKeys') or [] if k in roles_of}
        left = {roles_of[k][1] for k in prog.get('remainingRoleKeys') or [] if k in roles_of}
        self.batch_plan = {'batchId': b['batchId'], 'execution': b.get('execution'), 'companies': plan, 'done': done, 'left': left,
                           'oldState': b.get('state')}
        self.report['batch'] = {'batchId': b['batchId'], 'oldState': b.get('state'), 'newState': '等待正向评审（须用户重新确认一次）',
                                'companies': [f"{p['name']}（{p['code']}）：{'、'.join(p['roles'])}" for p in plan],
                                'processed': sorted(done), 'pending': sorted(left)}

    def find_role(self, j):
        for name, s, idx, row in S.iter_status_rows(self.ctx, self.master):
            if S.cell_value(row, idx.get(M.ROLE_CODE_COL, 0)) == j:
                return name, row, idx
        return None

    # ---------- 投递组件私有 ----------
    def convert_private(self):
        rb, unmatched = {}, []
        for key, roles in self.role_bindings.items():
            company = self.legacy_company.get(key, key)
            rb[key] = {}
            for rid, tpl in roles.items():
                j = self.role_ids.get((company, str(rid)))
                if not j:  # 例如申请页地址里的岗位编号：按已存投递链接对上
                    hits = {jj for (cc, _), jj in self.role_ids.items() if cc == company
                            for hit in [self.find_role(jj)] if hit and str(rid) in str(S.cell_value(hit[1], hit[2][M.APPLY_COL]) or '')}
                    j = next(iter(hits)) if len(hits) == 1 else None
                if not j:
                    unmatched.append(f'{key}/{rid} → {tpl}')
                    continue
                if j in rb[key] and rb[key][j] != tpl:
                    self.err(f'{key} 的两个旧岗位ID 对上同一岗位 {j} 却绑定了不同模板')
                rb[key][j] = tpl
        for u in unmatched:
            self.err(f'模板绑定对不上岗位：{u}')
        self.new_role_bindings = rb
        unknown = [k for k in list(self.company_bindings) + list(self.role_bindings) if k not in self.site_keys.values()]
        for k in unknown:
            self.err(f'模板绑定的目录项 {k} 不是任何站点的资产键')
        self.report['private'] = {
            'roleBindings': {k: v for k, v in rb.items()},
            'companyBindingsUnchanged': self.company_bindings,
            'channelOwnersReleased': self.owners,
            'queueFilesDiscarded': len(self.queue_files),
        }

    # ---------- 校验 ----------
    def validate(self):
        probs = (S.validate_codes(self.ctx, self.master) + S.validate_labels(self.ctx, self.master)
                 + S.validate_fields(self.ctx, self.master) + S.validate_link_exceptions(self.ctx, self.master))
        for p in probs:
            self.err('迁移后门禁：' + p)
        companies, sites = M.companies(self.master), M.sites(self.master)
        self.report['counts'] = {'companies': len(companies), 'sites': len(sites),
                                 'roles': sum(1 for n, s, idx, row in S.iter_status_rows(self.ctx, self.master) if S.row_kind(row, idx) == 'role'),
                                 'companyFiles': len(self.firms), 'counters': self.master.get(M.COUNTERS)}

    def run(self, quote=None):
        at = S.now()
        self.load()
        self.assign()
        self.move_role_values(quote or '（迁移报告确认后填入用户原话）', at)
        self.convert_ledger()
        self.convert_batch()
        self.convert_private()
        self.validate()
        return self.report

    # ---------- 写 ----------
    def apply(self, quote, rollback):
        if self.report['errors']:
            raise SystemExit('迁移报告有错误，不能写入：\n' + '\n'.join(self.report['errors']))
        tracked = [str(self.handoff.relative_to(self.ws)), str(self.ctx.book.relative_to(self.ws))]
        st = subprocess.run(['git', '-C', str(self.ws), 'status', '--porcelain', '--', *tracked], capture_output=True, text=True)
        if st.returncode or st.stdout.strip():
            raise SystemExit('岗位交接/ 与 Excel 须先提交为 git 回滚点（当前有未提交改动或不是 git 仓库）：\n' + st.stdout)
        # gitignored 的私有数据先打包到工作区以外
        rollback = Path(rollback).resolve()
        if self.ws in rollback.parents or rollback == self.ws:
            raise SystemExit('回滚包必须放在工作区以外')
        rollback.mkdir(parents=True, exist_ok=True)
        pack = rollback / f'step1-{time.strftime("%Y%m%d-%H%M%S")}.tar.gz'
        with tarfile.open(pack, 'w:gz') as t:
            for p in [self.ctx.ledger_path, self.ctx.labels_path, self.private / 'templates' / 'role-bindings.json',
                      self.private / 'runs' / 'company-channel-owners.json', *self.queue_files]:
                if p.exists():
                    t.add(p, arcname=str(p.relative_to(self.ws)))
        # 主表 + 台账
        S.save_master(self.ctx, self.master)
        S.save_json(self.ctx.ledger_path, self.ledger)
        # 交接宽表
        refreshed = self.store.refresh()
        if self.batch:
            bp = self.batch_plan
            docs = self.store.docs()
            for p in bp['companies']:
                d = docs[p['code']]
                roles = {r['岗位编码']: r for r in d['岗位']}
                chosen = sorted(p['roles'], key=lambda j: (roles[j]['志愿分组'], roles[j].get('志愿顺序') or 1))
                d['批次'] = {'批次编号': bp['batchId'], '状态': '等待正向评审', '公司顺序': p['order'], '选中岗位': chosen, '排除': [],
                            '执行': bp['execution'], '确认': None,
                            '进度': {'已处理': sorted(bp['done'] & set(chosen)), '待投递': sorted(bp['left'] & set(chosen))},
                            '更新时间': S.now()}
                self.store.save(d)
            self.store.review(bp['batchId'])
        # Excel 受控首写
        written = S.write(self.ctx, code_migration=True)
        # 旧交接文件
        removed = []
        for firm in self.firms:
            firm['file'].unlink()
            removed.append(str(firm['file'].relative_to(self.ws)))
            if not any(firm['file'].parent.iterdir()):
                firm['file'].parent.rmdir()
        for region in [p for p in self.base.iterdir() if p.is_dir()]:
            if not any(region.iterdir()):
                region.rmdir()
        for name in ('投递批次.json',):
            f = self.handoff / name
            if f.exists():
                f.unlink()
                removed.append(str(f.relative_to(self.ws)))
        # 投递组件私有
        tpl = self.private / 'templates'
        S.save_json(tpl / 'site-asset-keys.json', dict(sorted(self.site_keys.items())))
        S.save_json(tpl / 'role-bindings.json', self.new_role_bindings)
        if (self.private / 'runs' / 'company-channel-owners.json').exists():
            S.save_json(self.private / 'runs' / 'company-channel-owners.json', {})
        for q in self.queue_files:
            q.unlink()
        # 标签配置：投递批次文件已不存在
        cfg = json.loads(self.ctx.labels_path.read_text(encoding='utf-8'))
        cfg.get('companyStore', {}).pop('batchFile', None)
        S.save_json(self.ctx.labels_path, cfg)
        return {'rollbackPack': str(pack), 'companyFiles': refreshed, 'excel': written, 'removed': len(removed)}


def text_report(r):
    lines = ['# 第 1 步编码迁移报告', '']
    lines.append(f"主表与 Excel 迁移前一致：{'是' if r.get('excelInSyncBeforeMigration') else '否'}")
    c = r.get('counts') or {}
    a = r.get('assigned') or {}
    lines.append(f"分配编码：公司 {a.get('C')} 个、站点 {a.get('S')} 个、岗位 {a.get('J')} 个；迁移后共公司 {c.get('companies')}、站点 {c.get('sites')}、岗位行 {c.get('roles')}")
    lines.append(f"来自公司文件夹/登录配置、主表没有岗位的站点：{len(r.get('sitesFromCompanyFolders') or [])} 个")
    lines += ['  - ' + x for x in r.get('sitesFromCompanyFolders') or []]
    lines.append(f"公司.json 搬进主表的值：{len(r.get('roleValuesMoved') or [])} 项")
    lines += ['  - ' + x for x in r.get('roleValuesMoved') or []]
    lg = r.get('ledger') or {}
    for k in ('entries', 'link_exceptions'):
        v = lg.get(k) or {}
        lines.append(f"台账 {k}：换成编码 {v.get('converted')} 条；对不上现有行的已完成旧记录 {v.get('droppedFinishedHistory')} 条（不再保留）；未完成却对不上 {len(v.get('unresolved') or [])} 条")
    lines.append(f"删除 identity_changes：{lg.get('identityChangesRemoved')} 条")
    b = r.get('batch')
    if b:
        lines.append(f"当前批次 {b['batchId']}（原状态 {b['oldState']}）→ 各家「批次」，状态 {b['newState']}")
        lines += ['  - ' + x for x in b['companies']]
        lines.append(f"  已处理：{'、'.join(b['processed']) or '无'}；待投递：{'、'.join(b['pending']) or '无'}")
    p = r.get('private') or {}
    lines.append(f"站点资产键绑定：{len(r.get('siteAssetKeys') or {})} 条（旧目录项编号保持不变，表单资产不受影响）")
    lines.append('模板绑定（岗位ID → 岗位编码）：' + json.dumps(p.get('roleBindings'), ensure_ascii=False))
    lines.append(f"作废运行队列 {p.get('queueFilesDiscarded')} 个文件；释放公司批次账号占用：{json.dumps(p.get('channelOwnersReleased'), ensure_ascii=False)}")
    lines.append('')
    lines.append(f"警告 {len(r['warnings'])} 条：")
    lines += ['  - ' + x for x in r['warnings']]
    lines.append(f"错误 {len(r['errors'])} 条：")
    lines += ['  - ' + x for x in r['errors']]
    return '\n'.join(lines)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--workspace', required=True)
    ap.add_argument('--labels', default='产品管理/投递状态标签.json')
    ap.add_argument('--private-dir', default='.runtime/application')
    ap.add_argument('--queue-dir', default='.runtime/role-queue')
    ap.add_argument('--report', help='把报告 JSON 写到这里（工作区以外）')
    ap.add_argument('--apply', action='store_true')
    ap.add_argument('--quote', help='用户确认迁移报告的原话（--apply 必填）')
    ap.add_argument('--rollback-dir', help='gitignored 私有数据的回滚包目录（工作区以外，--apply 必填）')
    a = ap.parse_args()
    if a.apply and (not (a.quote or '').strip() or not a.rollback_dir):
        raise SystemExit('--apply 必须带 --quote "<用户原话>" 和 --rollback-dir')
    m = Migration(a.workspace, a.labels, a.private_dir, a.queue_dir)
    report = m.run(a.quote.strip() if a.quote else None)
    if a.report:
        Path(a.report).write_text(json.dumps(report, ensure_ascii=False, indent=2, default=sorted), encoding='utf-8')
    print(text_report(report))
    if a.apply:
        print(json.dumps(m.apply(a.quote.strip(), a.rollback_dir), ensure_ascii=False, indent=2))
    elif report['errors']:
        sys.exit(2)


if __name__ == '__main__':
    main()
