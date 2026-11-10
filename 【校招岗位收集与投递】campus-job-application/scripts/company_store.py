#!/usr/bin/env python3
"""交接区（中间区域）：岗位交接/公司/<公司编码>.json，每家公司一份宽表文件，只由本脚本从主表生成。

记录唯一、覆盖更新：
- 「站点」「岗位」两部分是主表的投影（公司主数据、投递站点、岗位表），每次 refresh 按主表覆盖重写，
  不在这里改；改值只能改主表（json_excel_sync.py 的固定命令），再 refresh。
- 「批次」部分是本轮交接的选岗、确认与进度，只在这里：forward 写入，forward-confirm 记用户原话与内容指纹，
  reverse / reverse-confirm 推进。
- 内容指纹只覆盖本批选中岗位的执行内容（链接、站点、志愿、城市、材料）和它们所在站点，用来核对“确认后有没有被改动”；
  它只当字段，不当键，也不用来查找。
- 链接只在投递端“打开失败”“第一次拿到空着的链接”或“岗位链接跳转后岗位编号变了（岗位名一致）”时，
  经投递回传 → apply-observations → 用户确认 → 主表 → refresh 刷新；编号不变的跳转不算变化。

子命令：
  refresh             按主表重建全部公司宽表文件（保留各自的「批次」部分）
  set-order           记录用户给定的志愿顺序（分组内不重复的正整数），写主表
  forward             正向导出：主表“岗位待投递”岗位 → 各公司文件的「批次」（等待正向评审）+ 评审页
  forward-confirm     用户评审确认后逐家登记原话与内容指纹，批次 → 已交接
  amend-materials     只补材料：预览 → 用户确认 → 更新材料并按“材料”作用域重新确认
  apply-observations  投递回传里的链接观察：预览 → 用户确认 → 写主表 → refresh；在批次中的公司按“链接”作用域重新确认
  reverse             反向评审：读取投递回传，生成“当前→更新”与反向评审页
  reverse-confirm     用户确认后写主表状态、清除已处理的回传；全部处理完 → 已回写
  batch-status        只读：显示本轮各公司的批次状态
"""
import argparse
import collections
import hashlib
import html
import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlparse

sys.path.insert(0, str(Path(__file__).resolve().parent))
import json_excel_sync as S  # noqa: E402
import master_codes as M  # noqa: E402

SCHEMA = 3
ACTIVE = ('已交接', '进行中', '等待反向评审')
HANDED_OFF = ('已交接', '进行中')
NOT_HANDED = '不交接'
FINISHED = (None, '已回写', NOT_HANDED)
LIST_PAGE_PROBLEM = '岗位链接是岗位列表页，不是具体岗位页'
LIST_PAGE = re.compile(r'(^|/)(position|job|jobs|post)s?/list(/|$)|list\.html$|/campus-recruit-list', re.I)
# 内容指纹覆盖的执行字段（投递组件按同一清单核对）
EXEC_ROLE_FIELDS = ('岗位编码', '站点编码', '岗位名称', '岗位链接', '投递链接', '官方岗位ID', 'BU', '志愿分组', '志愿顺序', '城市选项', '链接说明', '材料')
EXEC_SITE_FIELDS = ('站点编码', '站点', '登录链接', '账号核对页')
OBSERVATION_FIELDS = {'role': (S.LINK_COL, M.APPLY_COL), 'site': ('登录链接', '账号核对页')}


def now():
    return datetime.now(timezone.utc).astimezone().isoformat(timespec='seconds')


def sha(b):
    return hashlib.sha256(b).hexdigest()


def pyjson(v):
    """与投递组件 pyJson 逐字节一致。"""
    return json.dumps(v, ensure_ascii=False, sort_keys=True)


MATERIAL_DEFINITIONS_FILE = Path(__file__).resolve().parents[1] / 'references' / 'role-materials.json'


def material_definitions():
    """岗位材料定义只在 references/role-materials.json；新增材料只加一项。"""
    return json.loads(MATERIAL_DEFINITIONS_FILE.read_text(encoding='utf-8'))['materials']


def role_materials(values, role_code):
    """岗位材料只来自主表这一行的对应列；证据记岗位编码、列名和该单元格内容的指纹。"""
    items = []
    for d in material_definitions():
        raw = str(values.get(d['column']) or '').strip()
        first = re.split(d['separator'], raw)[0].strip()
        absent = not raw or bool(re.match(d['absent'], raw))
        value = first if not absent and re.fullmatch(d['value'], first) else None
        status = 'available' if value else 'absent' if absent else 'unresolved'
        items.append({'key': d['key'], 'status': status, 'value': value, 'raw': raw,
                      'requireWhenAvailable': d['requireWhenAvailable'],
                      'evidence': {'岗位编码': role_code, 'column': d['column'], 'cellSha256': sha(raw.encode())}})
    return {'schemaVersion': 2, 'items': items}


def handoff_view(doc):
    """内容指纹覆盖的范围：本批选中岗位的执行字段 + 它们所在站点。投递组件 handoffView 与此一致。"""
    batch = doc.get('批次') or {}
    chosen = set(batch.get('选中岗位') or [])
    roles = [{f: r.get(f) for f in EXEC_ROLE_FIELDS} for r in doc.get('岗位', []) if r['岗位编码'] in chosen]
    used = {r['站点编码'] for r in roles}
    sites = [{f: s.get(f) for f in EXEC_SITE_FIELDS} for s in doc.get('站点', []) if s['站点编码'] in used]
    return {'公司编码': doc['公司编码'], '批次编号': batch.get('批次编号'),
            '岗位': sorted(roles, key=lambda r: r['岗位编码']), '站点': sorted(sites, key=lambda s: s['站点编码'])}


def content_sha(doc):
    return sha(pyjson(handoff_view(doc)).encode())


SCOPE_LABEL = {'selection': 'selection', 'materials': 'materials_only', 'links': 'links_only'}


def set_confirmation(doc, scope, quote, at, **extra):
    """确认按作用域各留最新一条；顶层是当前生效的门禁（最近一次确认）。"""
    batch = doc['批次']
    prior = batch.get('确认') or {}
    scopes = dict(prior.get('scopes') or {})
    current = content_sha(doc)
    scopes[scope] = {'userQuote': quote, 'confirmedAt': at, 'contentSha256': current, **extra}
    batch['确认'] = {'userQuote': quote, 'confirmedAt': at, 'contentSha256': current,
                    'scope': SCOPE_LABEL[scope], **extra, 'scopes': scopes}


def verify_handoff(doc):
    """供 Hook/投递组件调用：该公司本批已交接且内容未被改动。"""
    batch = (doc or {}).get('批次') or {}
    if batch.get('状态') not in HANDED_OFF:
        return '投递批次未经用户正向评审确认'
    if (batch.get('确认') or {}).get('contentSha256') != content_sha(doc):
        return '投递批次在确认后被改动，须重新正向评审'
    return None


class Store:
    def __init__(self, workspace, labels):
        self.ctx = S.Ctx(workspace, labels)
        self.root = self.ctx.root
        cfg = self.ctx.cfg.get('companyStore', {})
        self.base = self.root / cfg.get('dir', '岗位交接/公司')
        self.handoff = self.base.parent
        self.outcome_file = self.handoff / '投递回传.json'
        self.third_party = [h.lower() for h in cfg.get('thirdPartyHosts', [])]

    # ---------- 读取 ----------
    def master_rows(self, master):
        """按公司编码分组的地区表行。"""
        out = collections.defaultdict(list)
        for name, s, idx, row in S.iter_status_rows(self.ctx, master):
            d = {h: S.cell_value(row, c) for h, c in idx.items()}
            out[d.get(M.COMPANY_CODE_COL)].append({'sheet': name, 'row': row, 'idx': idx, 'values': d,
                                                   'kind': S.row_kind(row, idx), 'key': S.row_key(name, row, idx)})
        return out

    def file(self, code):
        return self.base / f'{code}.json'

    def docs(self):
        out = {}
        for f in sorted(self.base.glob('C*.json')):
            if M.kind_of(f.stem) == 'C':
                out[f.stem] = json.loads(f.read_text(encoding='utf-8'))
        return out

    def save(self, doc):
        S.save_json(self.file(doc['公司编码']), doc)

    def round(self, docs=None):
        """本轮：处于交接流程中的批次编号及其公司文件（同一时刻只有一轮）。"""
        docs = self.docs() if docs is None else docs
        live = {}
        for code, d in docs.items():
            b = d.get('批次') or {}
            if b.get('批次编号') and b.get('状态') not in FINISHED:
                live.setdefault(b['批次编号'], {})[code] = d
        if len(live) > 1:
            raise SystemExit('交接区同时存在多个批次：' + '、'.join(sorted(live)) + '，须人工核对')
        return next(iter(live.items()), (None, {}))

    @staticmethod
    def role_id(url):
        """网站自己的岗位ID（值，不当键）：查询参数里的岗位ID，或最后一个含数字的路径片段；取不到即为 None。"""
        p = urlparse(str(url or ''))
        q = '&' + p.query + '&' + p.fragment
        for key in ('postId', 'jobAdId', 'jobId', 'jobCode', 'positionId', 'position_id', 'job_id', 'rid', 'pid', 'id'):
            m = re.search(r'(?:^|[?&#/])' + key + r'=([0-9A-Za-z\-]{3,})', q, re.I) \
                or re.search(r'/' + key + r'/([0-9A-Za-z\-]{3,})(?:\.html?)?(?:/|$)', p.path, re.I)
            if m and re.search(r'\d', m.group(1)):
                return m.group(1)
        segs = [re.sub(r'\.(html?|aspx?|php)$', '', x) for x in re.split(r'[/#?&=]', p.path + '/' + p.fragment)]
        segs = [x for x in segs if len(x) >= 4 and re.search(r'\d', x) and re.fullmatch(r'[0-9A-Za-z\-_]+', x)]
        return segs[-1] if segs else None

    def is_third_party(self, url):
        host = urlparse(str(url or '')).netloc.lower()
        return any(host == h or host.endswith('.' + h) for h in self.third_party)

    def link_problems(self, link, rid, dup):
        """写入时的链接质量规则（不是运行时检查）。"""
        p = urlparse(str(link or ''))
        out = []
        if p.scheme != 'https' or not p.netloc:
            out.append('岗位链接不是 https 官网链接')
        if self.is_third_party(link):
            out.append(f'岗位链接指向第三方招聘网站（{p.netloc}），需换成官网链接')
        if LIST_PAGE.search(p.path + ('#' + p.fragment if p.fragment else '')):
            out.append(LIST_PAGE_PROBLEM)
        if not rid:
            out.append('岗位链接里没有官方岗位ID（首页或列表页），不是具体岗位页')
        if dup:
            out.append('与同公司其他岗位共用同一链接，不是具体岗位页')
        return out

    # ---------- 志愿分组 ----------
    def bu_parallel(self, master):
        return {v.get(M.COMPANY_CODE_COL): v.get('是否分BU可平行投递') for _, v in S.quota_records(master) if v.get('范围') == '公司总量'}

    @staticmethod
    def group_of(parallel, bu):
        bu = str(bu or '').strip()
        return bu if parallel == '是' and bu and bu != '未披露' else '公司'

    @staticmethod
    def order_problems(roles):
        """同一志愿分组内：只有一个岗位时可不填；多个岗位时每个都必须有用户给定的、不重复的正整数顺序。"""
        groups = collections.defaultdict(list)
        for r in roles:
            groups[r.get('志愿分组') or '公司'].append(r.get('志愿顺序'))
        bad = {}
        for g, orders in groups.items():
            if len(orders) == 1 and orders[0] in (None, ''):
                continue
            if not all(isinstance(o, int) and not isinstance(o, bool) and o > 0 for o in orders) \
                    or len(set(orders)) != len(orders):
                bad[g] = orders
        return bad

    # ---------- 宽表投影 ----------
    def project(self, master, code, records):
        """一家公司的「站点」「岗位」投影：全部来自主表。"""
        ledger = self.ctx.ledger()
        parallel = self.bu_parallel(master).get(code)
        sites = [{'站点编码': sc, **{k: v[k] for k in ('站点', '登录链接', '账号核对页', '状态')}}
                 for sc, v in sorted(M.sites(master).items()) if v['公司编码'] == code]
        roles = []
        for r in records:
            if r['kind'] != 'role':
                continue
            v = r['values']
            link = v.get(S.LINK_COL)
            note, confirmed = S.link_exception_note(self.ctx, ledger, r['sheet'], r['row'], r['idx'])
            cities = [x.strip() for x in str(v.get(M.CITY_COL) or '').split(M.ALIAS_SEP) if x.strip()]
            role = {'岗位编码': v.get(M.ROLE_CODE_COL), '地区': r['sheet'], '站点编码': v.get(M.SITE_CODE_COL),
                    '岗位名称': v.get(S.NAME_COL), '岗位链接': link, '投递链接': v.get(M.APPLY_COL),
                    '官方岗位ID': self.role_id(link), 'BU': v.get('所属BU') or '未披露',
                    '志愿分组': self.group_of(parallel, v.get('所属BU')), '志愿顺序': v.get(M.ORDER_COL),
                    '城市选项': cities, '投递状态': v.get(self.ctx.column) or '',
                    '链接说明': note[len(S.LINK_EXCEPTION_PREFIX):] if note and confirmed else None,
                    '材料': role_materials(v, v.get(M.ROLE_CODE_COL))}
            roles.append(role)
        return sites, roles

    def build(self, master, code, records, old=None):
        sites, roles = self.project(master, code, records)
        return {'schemaVersion': SCHEMA, '公司编码': code, '公司名称': M.company_name(master, code),
                '志愿分组依据': '投递名额“公司总量”行 是否分BU可平行投递=' + str(self.bu_parallel(master).get(code) or '无名额规则'),
                '站点': sites, '岗位': roles, '批次': (old or {}).get('批次'),
                '生成自': {'主表指纹': master.get('data_sha256'), '时间': now()}}

    # ---------- refresh ----------
    def refresh(self, codes=None):
        """按主表覆盖重建公司宽表文件（保留「批次」）。codes 给定时只重建这些公司。
        主表里已没有岗位也没有站点的公司：删除其文件；但处于交接流程中的文件不删、报错由人工处理。"""
        master = self.ctx.load_master()
        probs = S.validate_codes(self.ctx, master)
        if probs:
            raise SystemExit('主表编码门禁未通过，不能生成交接文件：\n' + '\n'.join(probs[:20]))
        rows = self.master_rows(master)
        site_owners = {v['公司编码'] for v in M.sites(master).values()}
        wanted = {c for c, rs in rows.items() if c and any(r['kind'] == 'role' for r in rs)} | site_owners
        docs = self.docs()
        written, removed = 0, []
        for code in sorted(wanted if codes is None else set(codes)):
            if code not in wanted:
                continue
            doc = self.build(master, code, rows.get(code, []), docs.get(code))
            self.save(doc)
            written += 1
        if codes is None:
            for code, d in docs.items():
                if code in wanted:
                    continue
                if (d.get('批次') or {}).get('状态') in ACTIVE:
                    raise SystemExit(f'{code} 已不在主表，但它的批次仍在交接流程中，须人工核对')
                self.file(code).unlink()
                removed.append(code)
        return {'written': written, 'removed': removed}

    # ---------- set-order ----------
    def set_order(self, items, quote):
        """用户给定志愿顺序：[{code: 岗位编码, 志愿顺序}]；同公司同分组一次给全。
        本次涉及的分组里没给顺序的岗位清掉旧顺序，避免新旧混用。"""
        if not quote or not quote.strip():
            raise SystemExit('必须提供用户给定顺序的原话')
        master = self.ctx.load_master()
        rows = self.master_rows(master)
        by_code = {}
        for code, rs in rows.items():
            for r in rs:
                if r['kind'] == 'role':
                    by_code[r['values'].get(M.ROLE_CODE_COL)] = (code, r)
        parallel = self.bu_parallel(master)
        given = {}
        for it in items:
            o = it.get('志愿顺序', it.get('顺序'))
            if not isinstance(o, int) or isinstance(o, bool) or o < 1:
                raise SystemExit(f'顺序必须是正整数：{it}')
            if it.get('code') not in by_code:
                raise SystemExit(f"岗位编码不存在：{it.get('code')}")
            given[it['code']] = o
        touched = collections.defaultdict(set)
        for j in given:
            c, r = by_code[j]
            touched[c].add(self.group_of(parallel.get(c), r['values'].get('所属BU')))
        plan, out = [], []
        for c, groups in touched.items():
            members = [(r['values'][M.ROLE_CODE_COL], self.group_of(parallel.get(c), r['values'].get('所属BU')), r)
                       for r in rows[c] if r['kind'] == 'role']
            members = [m for m in members if m[1] in groups]
            check = [{'志愿分组': g, '志愿顺序': given.get(j)} for j, g, _ in members if j in given]
            bad = self.order_problems(check)
            if bad:
                raise SystemExit(f'{M.company_name(master, c)} 志愿顺序不合格（分组内须为不重复的正整数）：{bad}')
            for j, g, r in members:
                if given.get(j) != r['values'].get(M.ORDER_COL):
                    plan.append({'code': j, '志愿顺序': given.get(j)})
            out.append({'公司': M.company_name(master, c), '分组': {g: sorted((given[j], r['values'].get(S.NAME_COL)) for j, gg, r in members if gg == g and j in given)
                                                                   for g in sorted(groups)}})
        if plan:
            S.set_order(self.ctx, plan, quote)
        self.refresh(sorted(touched))
        return out

    # ---------- forward ----------
    def forward(self, session=None, turn=None):
        from application_requirements import Requirements
        return Requirements(self.root, session, turn).build(self._forward)

    def _forward(self, requirements):
        from application_requirements import select_companies
        live, _ = self.round()
        docs = self.docs()
        for code, d in docs.items():
            if ((d.get('批次') or {}).get('状态')) in ACTIVE:
                raise SystemExit(f"当前批次 {live} 仍在进行（{d['公司名称']}：{d['批次']['状态']}），只有全部 已回写 后才能生成新批次")
        master = self.ctx.load_master()
        rows = self.master_rows(master)
        names = {c: [v['名称']] + v['别名'] for c, v in M.companies(master).items()}
        ledger = self.ctx.ledger()
        chosen_codes = select_companies(rows, ledger, requirements, names)  # 选不出公司时在写任何文件之前失败
        self.refresh()
        batch_id = 'B' + datetime.now().strftime('%Y%m%d%H%M%S')
        browser_channels = None
        if requirements.get('浏览器'):
            browser_channels = ['chrome', 'iab'] if requirements['浏览器'] == 'dual' else [requirements['浏览器']]
        docs = self.docs()
        for code, d in docs.items():  # 只留本轮：其它公司的旧批次（已回写）清掉
            if code not in chosen_codes and d.get('批次'):
                d['批次'] = None
                self.save(d)
        summary = {'companies': 0, 'roles': 0, 'excluded': 0}
        for order, code in enumerate(chosen_codes, start=1):
            doc = docs[code]
            records = rows[code]
            pending = [r for r in records if r['kind'] == 'role' and r['values'].get(self.ctx.column) == '岗位待投递'
                       and (not requirements.get('地区') or r['sheet'] == requirements['地区'])]
            by_code = {r['岗位编码']: r for r in doc['岗位']}
            links = collections.Counter(str(r['values'].get(S.LINK_COL) or '').rstrip('/') for r in records if r['kind'] == 'role')
            ok, excluded = [], []
            for r in pending:
                j = r['values'].get(M.ROLE_CODE_COL)
                role = by_code[j]
                link = role['岗位链接']
                problems = self.link_problems(link, role['官方岗位ID'], links[str(link or '').rstrip('/')] > 1)
                note, confirmed = S.link_exception_note(self.ctx, ledger, r['sheet'], r['row'], r['idx'])
                if note and not confirmed:
                    problems.append('备注是链接例外，但没有用户确认记录，须用 link-exception 登记')
                elif note:
                    problems = [x for x in problems if x != LIST_PAGE_PROBLEM]
                if problems:
                    excluded.append({'岗位编码': j, '岗位': role['岗位名称'], '原因': '；'.join(problems)})
                    continue
                ok.append(role)
            bad = self.order_problems(ok)
            for g, orders in bad.items():
                for role in [x for x in ok if x['志愿分组'] == g]:
                    excluded.append({'岗位编码': role['岗位编码'], '岗位': role['岗位名称'],
                                     '原因': f'志愿分组「{g}」有 {len(orders)} 个待投递岗位，志愿顺序未由用户给全（现为 {orders}）；用 set-order 记录用户给定的顺序后重新生成'})
            ok = sorted([x for x in ok if x['志愿分组'] not in bad], key=lambda x: (x['志愿分组'], x.get('志愿顺序') or 1))
            doc['批次'] = {'批次编号': batch_id, '状态': '等待正向评审' if ok else NOT_HANDED, '公司顺序': order,
                         '选中岗位': [x['岗位编码'] for x in ok], '排除': excluded,
                         '确认': None, '进度': None, '更新时间': now()}
            self.save(doc)
            summary['companies'] += bool(ok)
            summary['roles'] += len(ok)
            summary['excluded'] += len(excluded)
        summary.update(self.forward_summary(batch_id))
        from browser_capabilities import application_policy_file
        from application_requirements import save as save_execution_policy
        save_execution_policy(application_policy_file(self.root), {'batchId': batch_id, 'browserChannels': browser_channels, 'authorizationRef': 'forward_requirements'})
        return {'batchId': batch_id, **summary, 'browserChannels': browser_channels}

    def _fresh_view_matches(self, master, rows, doc):
        """确认前核对：用当前主表重建的执行内容与文件里的一致（用户看到的就是要确认的）。"""
        fresh = self.build(master, doc['公司编码'], rows.get(doc['公司编码'], []), doc)
        return handoff_view(fresh) == handoff_view(doc)

    def forward_confirm(self, quote):
        if not quote or not quote.strip():
            raise SystemExit('必须提供用户确认原话')
        batch_id, docs = self.round()
        waiting = {c: d for c, d in docs.items() if d['批次']['状态'] == '等待正向评审'}
        if not waiting:
            raise SystemExit('没有等待正向评审的批次')
        master = self.ctx.load_master()
        rows = self.master_rows(master)
        for code, d in waiting.items():
            if not self._fresh_view_matches(master, rows, d):
                raise SystemExit(f"{d['公司名称']} 的交接内容与主表不一致（材料或链接在生成评审后变了），请重新生成正向评审")
        at = now()
        for d in waiting.values():
            d['批次']['确认'] = None
            set_confirmation(d, 'selection', quote.strip(), at)
            d['批次']['状态'], d['批次']['更新时间'] = '已交接', at
            self.save(d)
        return {'batchId': batch_id, 'state': '已交接', 'companies': sorted(waiting)}

    def amend_materials(self, quote=None, expected_sha=None):
        """只补材料：不重新选岗、不改顺序；预览 → 用户确认 → 更新并按“材料”作用域重新确认。"""
        batch_id, docs = self.round()
        live = {c: d for c, d in docs.items() if d['批次']['状态'] in HANDED_OFF}
        if not live or any(verify_handoff(d) for d in live.values()):
            raise SystemExit('必须是已确认且未变更的当前批次')
        prior = combined_sha(live)
        master = self.ctx.load_master()
        rows = self.master_rows(master)
        changes, updated = [], {}
        for code, d in live.items():
            fresh = {r['岗位编码']: r for r in self.build(master, code, rows.get(code, []), d)['岗位']}
            for role in d['岗位']:
                if role['岗位编码'] in d['批次']['选中岗位'] and fresh[role['岗位编码']]['材料'] != role['材料']:
                    changes.append({'公司': d['公司名称'], '岗位': role['岗位名称'], 'current': role['材料'], 'updated': fresh[role['岗位编码']]['材料']})
                    role['材料'] = fresh[role['岗位编码']]['材料']
                    updated[code] = d
        preview = {'batchId': batch_id, 'previousContentSha256': prior, 'changes': changes, 'selectionChanged': False}
        if quote is None:
            return preview
        if not quote.strip() or expected_sha != prior:
            raise SystemExit('必须附用户授权原话和预览时的批次摘要')
        at = now()
        for code, d in updated.items():
            old = d['批次']['确认']['contentSha256']
            set_confirmation(d, 'materials', quote.strip(), at, previousContentSha256=old)
            d['批次']['更新时间'] = at
            self.save(d)
            if verify_handoff(json.loads(self.file(code).read_text(encoding='utf-8'))):
                raise ValueError('补充材料后批次确认读回失败')
        return preview

    def forward_summary(self, batch_id):
        """将当前批次清单直接返回给调用方。"""
        docs = [d for d in self.docs().values() if (d.get('批次') or {}).get('批次编号') == batch_id]
        companies, exclusions = [], []
        for doc in sorted(docs, key=lambda d: d['批次']['公司顺序']):
            company = {'公司编码': doc['公司编码'], '公司名称': doc['公司名称']}
            batch = doc['批次']
            by_code = {r['岗位编码']: r for r in doc['岗位']}
            roles = [{k: by_code[j][k] for k in ('岗位编码', '岗位名称', '志愿分组', '志愿顺序', '链接说明')}
                     for j in batch['选中岗位']]
            if roles:
                companies.append({**company, '状态': batch['状态'], '岗位': roles})
            exclusions.extend({**company, **item} for item in batch.get('排除') or [])
        return {'handoffCompanies': companies, 'excludedRoles': exclusions}

    # ---------- 投递回传 ----------
    def outcomes(self):
        f = self.outcome_file
        return json.loads(f.read_text(encoding='utf-8')) if f.exists() else {'schemaVersion': 2, 'roles': {}, 'observations': []}

    def save_outcomes(self, out):
        if out.get('roles') or out.get('observations'):
            S.save_json(self.outcome_file, {'schemaVersion': 2, 'roles': out.get('roles') or {}, 'observations': out.get('observations') or []})
        elif self.outcome_file.exists():
            self.outcome_file.unlink()

    def apply_observations(self, quote=None, expected_sha=None, replacements=()):
        """投递回传里的链接观察 → 预览 → 用户确认 → 写主表 → refresh。
        观察：{kind: role|site, code, field, value(null=打不开), evidence, observedAt}。
        岗位链接打不开时投递端给不出新值，须搜索端找到官网新链接后经 replacements=[{code, new_link}] 一并确认；
        岗位链接带值 = 跳转后编号变了的岗位页，直接作为新链接预览。"""
        out = self.outcomes()
        obs = out.get('observations') or []
        repl = {r['code']: r['new_link'] for r in replacements}
        relinks, applies, sites, waiting, lines = [], [], {}, [], []
        for o in obs:
            if o.get('field') not in OBSERVATION_FIELDS.get(o.get('kind'), ()):
                raise SystemExit(f'投递回传里有无法识别的观察：{o}')
            if o['kind'] == 'role' and o['field'] == S.LINK_COL:
                if o['code'] in repl:
                    relinks.append({'code': o['code'], 'new_link': repl[o['code']]})
                elif o.get('value'):
                    # 投递端打开岗位链接后跳到了编号不同的岗位页、且岗位名核对一致：观察值就是跳转后的岗位页，
                    # 经用户确认后换成它（官方岗位ID随新链接重算）。
                    relinks.append({'code': o['code'], 'new_link': o['value']})
                else:
                    waiting.append(o)
            elif o['kind'] == 'role':
                applies.append({'code': o['code'], M.APPLY_COL: o.get('value')})
            else:
                sites.setdefault(o['code'], {'code': o['code']})[o['field']] = o.get('value')
        for code in repl:
            if not any(r['code'] == code for r in relinks):
                relinks.append({'code': code, 'new_link': repl[code]})
        if relinks:
            lines += S.relink(self.ctx, relinks, None, preview=True, allow_in_batch=True)['show_user_exactly']
        if applies:
            lines += S.set_apply_link(self.ctx, applies, None, preview=True, allow_in_batch=True)['show_user_exactly']
        if sites:
            lines += S.set_site_links(self.ctx, list(sites.values()), None, preview=True, allow_in_batch=True)['show_user_exactly']
        info = [f"{o['code']} 岗位链接打不开（{o.get('evidence') or '无证据说明'}，{o.get('observedAt', '')}）：需要搜索端找到官网上的新链接，再随本命令的 --input 一并确认"
                for o in waiting]
        master = self.ctx.load_master()
        role_company = {r['values'].get(M.ROLE_CODE_COL): c for c, rs in self.master_rows(master).items() for r in rs if r['kind'] == 'role'}
        site_company = {sc: v['公司编码'] for sc, v in M.sites(master).items()}
        affected = sorted({role_company.get(x['code']) for x in relinks + applies} | {site_company.get(c) for c in sites} - {None})
        batch_id, docs = self.round()
        live = {c: d for c, d in docs.items() if c in affected and d['批次']['状态'] in ACTIVE}
        prior = combined_sha(live) if live else None
        for d in live.values():
            lines.append(f"批次 {batch_id}：{d['公司名称']} 的交接内容随之更新，按“链接”作用域重新确认；执行端只需重新导入这一家")
        if quote is None:
            return {'show_user_exactly': lines, 'waiting_for_new_link': info, 'previousContentSha256': prior,
                    'rule': '原样展示给用户并取得明确确认后，带 --quote（批次中有公司受影响时再带 --expected-sha）执行'}
        if not quote.strip() or (prior and expected_sha != prior):
            raise SystemExit('必须附用户确认原话，以及预览时的批次摘要')
        if not lines:
            return {'show_user_exactly': [], 'waiting_for_new_link': info}
        if relinks:
            S.relink(self.ctx, relinks, quote.strip(), allow_in_batch=True)
        if applies:
            S.set_apply_link(self.ctx, applies, quote.strip(), allow_in_batch=True)
        if sites:
            S.set_site_links(self.ctx, list(sites.values()), quote.strip(), allow_in_batch=True)
        self.refresh(affected)
        at = now()
        for code in live:
            d = json.loads(self.file(code).read_text(encoding='utf-8'))
            old = (d['批次'].get('确认') or {}).get('contentSha256')
            set_confirmation(d, 'links', quote.strip(), at, previousContentSha256=old)
            d['批次']['更新时间'] = at
            self.save(d)
        done = {(r['code'], S.LINK_COL) for r in relinks} | {(a['code'], M.APPLY_COL) for a in applies} | \
               {(c, f) for c, v in sites.items() for f in v if f != 'code'}
        out['observations'] = [o for o in obs if (o['code'], o['field']) not in done]
        self.save_outcomes(out)
        return {'applied': len(done), 'companiesRefreshed': affected, 'show_user_exactly': lines, 'waiting_for_new_link': info}

    def reverse(self):
        """反向评审：读取投递回传，逐行给出“当前状态是X，更新状态是Y”，生成反向评审页；不改主表。"""
        batch_id, docs = self.round()
        live = {c: d for c, d in docs.items() if d['批次']['状态'] in ACTIVE}
        if not live:
            raise SystemExit('没有已交接的投递批次')
        items, notes, _ = self._reverse_items()
        master = self.ctx.load_master()
        plan = S.propose(self.ctx, [{k: v for k, v in i.items() if k != 'outcomeKey'} for i in items], master) if items else []
        lines = [p['text'] for p in plan]
        at = now()
        for d in live.values():
            d['批次']['状态'], d['批次']['更新时间'] = '等待反向评审', at
            self.save(d)
        page = f"""<!doctype html><html lang=zh><meta charset=utf-8><meta name=viewport content="width=device-width,initial-scale=1"><title>反向评审</title>
<style>body{{font:14px/1.5 -apple-system,"PingFang SC",sans-serif;margin:16px;max-width:1100px}}li{{margin:4px 0}}</style>
<h1>反向评审 · 批次 {html.escape(batch_id)}</h1><p>生成时间：{html.escape(at)}</p>
<h2>主表状态变更（需你确认）</h2><ul>{''.join('<li>' + html.escape(t) + '</li>' for t in lines) or '<li>无</li>'}</ul>
<h2>不改状态、需要你知道或处理的结果</h2><ul>{''.join('<li>' + html.escape(n['title'] + '：' + n['说明']) + '</li>' for n in notes) or '<li>无</li>'}</ul></html>"""
        (self.handoff / '反向评审.html').write_text(page, encoding='utf-8')
        return {'batchId': batch_id, 'state': '等待反向评审', 'show_user_exactly': lines,
                'notes': [n['title'] + '：' + n['说明'] for n in notes], 'review': '岗位交接/反向评审.html'}

    def _reverse_items(self):
        """投递回传按岗位编码对到主表行：已投递/已关闭 → 状态变更；不确定/受阻/跳过 → 只列出待用户处理。"""
        roles = self.outcomes().get('roles') or {}
        master = self.ctx.load_master()
        known = {r['values'].get(M.ROLE_CODE_COL) for rs in self.master_rows(master).values() for r in rs if r['kind'] == 'role'}
        items, notes = [], []
        for code, o in sorted(roles.items()):
            if code not in known:
                notes.append({'key': code, 'event': o['event'], 'title': o.get('roleTitle', code), '说明': '回传的岗位编码不在主表，需人工核对'})
                continue
            if o['event'] in REVERSE_STATUS:
                items.append({'code': code, 'to': REVERSE_STATUS[o['event']], 'outcomeKey': code})
            else:
                notes.append({'key': code, 'event': o['event'], 'title': o.get('roleTitle', code),
                              '说明': '投递结果不确定或受阻，处理前禁止重跑' if o['event'] in UNRESOLVED else '用户跳过，状态不变'})
        return items, notes, roles

    def reverse_confirm(self, quote, resolved=()):
        """用户确认后：写主表状态（经 json_excel_sync 确认台账），清除已处理的回传；某家全部处理完 → 该家 已回写。"""
        if not quote or not quote.strip():
            raise SystemExit('必须提供用户确认原话')
        batch_id, docs = self.round()
        live = {c: d for c, d in docs.items() if d['批次']['状态'] == '等待反向评审'}
        if not live:
            raise SystemExit('没有等待反向评审的批次，请先运行 reverse')
        for d in live.values():
            if (d['批次'].get('确认') or {}).get('contentSha256') != content_sha(d):
                raise SystemExit(f"{d['公司名称']} 的投递批次在确认后被改动，须重新正向评审")
        items, notes, roles = self._reverse_items()
        if items:
            S.confirm(self.ctx, [{k: v for k, v in i.items() if k != 'outcomeKey'} for i in items], quote.strip(), f'反向评审 {batch_id}', refresh_handoff=False)
        done = {i['outcomeKey'] for i in items} | {n['key'] for n in notes if n['event'] not in UNRESOLVED} | set(resolved)
        left = {k: v for k, v in roles.items() if k not in done}
        out = self.outcomes()
        out['roles'] = left
        self.save_outcomes(out)
        summary = self.refresh_confirmed_statuses(live, done, left, batch_id)
        return {'batchId': batch_id, 'companies': summary, 'statusChanges': len(items), 'unresolved': sorted(left)}

    def refresh_confirmed_statuses(self, codes, done=(), unresolved=None, batch_id=None):
        """从主表更新受影响公司的宽表及本轮执行进度。"""
        docs = self.docs()
        codes = sorted(set(codes) & docs.keys())
        if not codes:
            return {}
        master = self.ctx.load_master()
        rows = self.master_rows(master)
        unresolved = self.outcomes().get('roles', {}) if unresolved is None else unresolved
        at, summary, batches = now(), {}, set()
        for code in codes:
            doc = self.build(master, code, rows.get(code, []), docs[code])
            batch = doc.get('批次') or {}
            if batch.get('状态') in ACTIVE:
                batches.add(batch['批次编号'])
                selected = batch['选中岗位']
                status = {r['岗位编码']: r['投递状态'] for r in doc['岗位']}
                processed = set((batch.get('进度') or {}).get('已处理') or [])
                processed.update(j for j in selected if j in done or status.get(j) in ('岗位已投递', '岗位失效/关闭'))
                pending = [j for j in selected if j not in processed and status.get(j) in ('岗位待投递', '岗位待筛选', None, '')]
                batch['进度'] = {'已处理': sorted(processed), '待投递': pending}
                batch['状态'] = ('等待反向评审' if any(j in unresolved for j in selected)
                               else '进行中' if pending else '已回写')
                batch['更新时间'] = at
                summary[code] = batch['状态']
            self.save(doc)
        if (batch_id or batches) and self.round()[0] is None:
            from application_requirements import Requirements
            Requirements(self.root).complete(batch_id or next(iter(batches)))
        return summary

    def batch_status(self):
        batch_id, docs = self.round()
        return {'batchId': batch_id, 'companies': {c: {'公司': d['公司名称'], '状态': d['批次']['状态'],
                                                     '岗位': len(d['批次']['选中岗位']), 'handoffProblem': verify_handoff(d) if d['批次']['选中岗位'] else None}
                                                 for c, d in docs.items()}}


REVERSE_STATUS = {'submission_confirmed': '岗位已投递', 'role_closed': '岗位失效/关闭'}
UNRESOLVED = ('submission_uncertain', 'application_blocked')


def combined_sha(docs):
    """多家公司一起确认时的预览摘要：各家内容指纹按公司编码排序后再取指纹。"""
    return sha(pyjson(sorted((c, content_sha(d)) for c, d in docs.items())).encode())


def main():
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--workspace', default=os.environ.get('CAMPUS_RECRUITING_WORKSPACE', '.'))
    ap.add_argument('--labels', default=os.environ.get('CAMPUS_STATUS_LABELS', '产品管理/投递状态标签.json'))
    ap.add_argument('--session', help='当前宿主会话身份，默认使用宿主环境')
    ap.add_argument('--turn', help='当前宿主回合身份，默认读取需求Hook上下文')
    sub = ap.add_subparsers(dest='cmd', required=True)
    p = sub.add_parser('refresh')
    p.add_argument('--company', action='append', help='按公司编码从主表刷新宽表与已确认投递进度')
    sub.add_parser('forward')
    p = sub.add_parser('amend-materials')
    p.add_argument('--quote', help='不传则仅预览；传用户明确授权原话后只补当前岗位材料')
    p.add_argument('--expected-sha', help='预览返回的 previousContentSha256')
    p = sub.add_parser('apply-observations')
    p.add_argument('--input', help='UTF-8 JSON：[{code: 岗位编码, new_link}]，为打不开的岗位链接给出搜索端找到的官网新链接')
    p.add_argument('--quote', help='不传则仅预览；传用户确认原话后写主表并刷新交接文件')
    p.add_argument('--expected-sha', help='预览返回的 previousContentSha256（有批次中的公司受影响时必填）')
    p = sub.add_parser('set-order')
    p.add_argument('--input', required=True, help='UTF-8 JSON：[{code: 岗位编码, 志愿顺序}]，同公司同分组一次给全')
    p.add_argument('--quote', required=True, help='用户给定顺序的原话')
    p = sub.add_parser('forward-confirm')
    p.add_argument('--quote', required=True)
    sub.add_parser('batch-status')
    sub.add_parser('reverse')
    p = sub.add_parser('reverse-confirm')
    p.add_argument('--quote', required=True)
    p.add_argument('--resolved', action='append', default=[], help='用户已处理的不确定/受阻结果（岗位编码），确认后清除')
    a = ap.parse_args()
    st = Store(a.workspace, a.labels)
    if a.cmd == 'refresh':
        with S.exclusive(st.ctx):
            out = {'companies': st.refresh_confirmed_statuses(a.company)} if a.company else st.refresh()
    elif a.cmd == 'forward':
        with S.exclusive(st.ctx):
            out = st.forward(a.session, a.turn)
    elif a.cmd == 'amend-materials':
        with S.exclusive(st.ctx):
            out = st.amend_materials(a.quote, a.expected_sha)
    elif a.cmd == 'apply-observations':
        repl = json.loads(Path(a.input).read_text(encoding='utf-8')) if a.input else []
        with S.exclusive(st.ctx):
            out = st.apply_observations(a.quote, a.expected_sha, repl)
    elif a.cmd == 'set-order':
        with S.exclusive(st.ctx):
            out = st.set_order(json.loads(Path(a.input).read_text(encoding='utf-8')), a.quote)
    elif a.cmd == 'forward-confirm':
        with S.exclusive(st.ctx):
            out = st.forward_confirm(a.quote)
    elif a.cmd == 'reverse':
        with S.exclusive(st.ctx):
            out = st.reverse()
    elif a.cmd == 'reverse-confirm':
        with S.exclusive(st.ctx):
            out = st.reverse_confirm(a.quote, a.resolved)
    else:
        out = st.batch_status()
    print(json.dumps(out, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
