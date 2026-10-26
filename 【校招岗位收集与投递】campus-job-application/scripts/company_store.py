#!/usr/bin/env python3
"""公司文件夹（中间区域）：岗位交接/公司/<地区>/<公司>/公司.json。

记录唯一、覆盖更新：公司.json 只存主表没有的投递运行数据（companyId、官网域名、官方岗位ID、
投递链接、志愿顺序、城市选项）及每个岗位的志愿分组，每次覆盖写，不存研究记录、方案原文、状态副本或任何历史。
志愿分组：该公司“投递名额”表“公司总量”行的“是否分BU可平行投递”为“是”时按所属BU分组（BU未披露归“公司”），
否则整家公司一个分组；志愿顺序只在分组内排，由用户给定（set-order 记录用户原话）。
同目录 登录.json / 登录@<域名>.json 为各站点登录配置，表单定义.json / 表单定义@<域名>.json 为表单定义；
表单模板只在投递组件的私有运行目录，不在这里存副本。
岗位字段与投递状态只在主表；投递规则只在主表“投递名额”表。

子命令：
  refresh   按主表重建全部 公司.json（覆盖写）；没有岗位行也没有登录/表单资产的公司文件夹删除。
  set-order 记录用户给定的志愿顺序（分组内不重复的正整数），附用户原话；覆盖写到 公司.json。
  forward   正向导出：主表“岗位待投递”岗位 → 投递批次.json（状态：等待正向评审，覆盖写）+ 评审页。
  forward-confirm  用户评审确认后登记原话与内容摘要，批次状态 → 已交接。
  reverse          反向评审：读取投递组件写入的 投递回传.json，生成“当前→更新”与反向评审页（批次 → 等待反向评审）。
  reverse-confirm  用户确认后写主表状态、清除已处理的回传；全部处理完批次 → 已回写，才能生成下一批。
  batch-status     只读：显示批次状态。
"""
import argparse
import collections
import hashlib
import html
import json
import os
import re
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlparse

sys.path.insert(0, str(Path(__file__).resolve().parent))
import json_excel_sync as S  # noqa: E402

UNSAFE = re.compile(r'[\\/:*?"<>|]')
STAGES_OK_FOR_NEW_BATCH = {None, '空闲', '已回写', '等待正向评审'}
KEEP_ROLE_FIELDS = ('投递链接', '顺序', '城市选项')
LIST_PAGE_PROBLEM = '岗位链接是岗位列表页，不是具体岗位页'
LIST_PAGE = re.compile(r'(^|/)(position|job|jobs|post)s?/list(/|$)|list\.html$|/campus-recruit-list', re.I)


def now():
    return datetime.now(timezone.utc).astimezone().isoformat(timespec='seconds')


def sha(b):
    return hashlib.sha256(b).hexdigest()


def safe(name):
    return UNSAFE.sub('／', str(name)).strip() or '未命名'


def kid(key):
    return json.dumps(key, ensure_ascii=False, sort_keys=True)


MATERIAL_DEFINITIONS_FILE = Path(__file__).resolve().parents[1] / 'references' / 'role-materials.json'


def material_definitions():
    """岗位材料定义只在 references/role-materials.json；新增材料只加一项。"""
    return json.loads(MATERIAL_DEFINITIONS_FILE.read_text(encoding='utf-8'))['materials']


def _location_matches(spec, url, role_id):
    """与投递组件 transitionLocationMatches 同一规则：路径正则 + 查询参数（$roleId 按岗位取值）。"""
    from urllib.parse import urlparse, parse_qs
    try:
        u = urlparse(str(url or ''))
    except ValueError:
        return False
    hp, _, hq = u.fragment.partition('?')
    if not spec or not re.search(spec.get('pathnamePattern') or r'^$', u.path):
        return False
    if spec.get('hashPathPattern') and not re.search(spec['hashPathPattern'], hp):
        return False
    for mapping, query in ((spec.get('query'), u.query), (spec.get('hashQuery'), hq)):
        params = parse_qs(query, keep_blank_values=True)
        for k, v in (mapping or {}).items():
            want = role_id if v == '$roleId' else v
            if want is None or params.get(k) != [str(want)]:
                return False
    return True


def entry_link_problem(folder, item):
    """组批时核对：该公司声明式入口能否认出这个岗位（岗位链接或投递链接之一）。"""
    links = [item.get('岗位链接'), item.get('投递链接')]
    entries = []
    for f in sorted(Path(folder).glob('登录*.json')):
        try:
            e = json.loads(f.read_text(encoding='utf-8')).get('entry') or {}
        except (OSError, ValueError):
            continue
        if e.get('kind') == 'declarative_transition':
            entries.append(e)
    if not entries:
        return None
    for e in entries:
        for phase in ('detail', 'form'):
            spec = (e.get(phase) or {}).get('location')
            if any(_location_matches(spec, l, item.get('岗位ID')) for l in links if l):
                return None
    return '岗位链接与公司入口配置不匹配（入口认不出这个岗位）；若网站已改链接，先由投递端提出链接变更、经你确认后用 relink-observed 覆盖，或重新采集入口'


def role_materials(values, master_hash):
    """Freeze applicable role materials from their sole master row, not company.json."""
    items = []
    for d in material_definitions():
        raw = str(values.get(d['column']) or '').strip()
        first = re.split(d['separator'], raw)[0].strip()
        absent = not raw or bool(re.match(d['absent'], raw))
        value = first if not absent and re.fullmatch(d['value'], first) else None
        status = 'available' if value else 'absent' if absent else 'unresolved'
        items.append({'key': d['key'], 'status': status, 'value': value, 'raw': raw,
                      'requireWhenAvailable': d['requireWhenAvailable'],
                      'evidence': {'masterDataSha256': master_hash, 'company': values.get(S.COMPANY_COL),
                                   'roleUrl': values.get(S.LINK_COL), 'column': d['column']}})
    return {'schemaVersion': 1, 'items': items}


def material_summary(role):
    labels = {d['key']: d['label'] for d in material_definitions()}
    return '；'.join(labels.get(item['key'], item['key']) + '：' + (
        item['value'] + '（有则必填）' if item['status'] == 'available'
        else '无已记录' if item['status'] == 'absent' else '待核实')
        for item in role.get('materials', {}).get('items', []))


class Store:
    def __init__(self, workspace, labels):
        self.ctx = S.Ctx(workspace, labels)
        self.root = self.ctx.root
        cfg = self.ctx.cfg.get('companyStore', {})
        self.base = self.root / cfg.get('dir', '岗位交接/公司')
        self.batch_file = self.root / cfg.get('batchFile', '岗位交接/投递批次.json')
        self.review_file = self.root / cfg.get('forwardReview', '岗位交接/正向评审.html')
        self.third_party = [h.lower() for h in cfg.get('thirdPartyHosts', [])]

    # ---------- 读取 ----------
    def master_rows(self, master):
        out = collections.defaultdict(list)
        for name, s, idx, row in S.iter_status_rows(self.ctx, master):
            d = {h: S.cell_value(row, c) for h, c in idx.items()}
            out[(name, d.get(S.COMPANY_COL))].append({'row': row, 'idx': idx, 'values': d,
                                                      'kind': S.row_kind(row, idx), 'key': S.row_key(name, row, idx)})
        return out

    # ---------- 规则 ----------
    @staticmethod
    def role_id(url):
        """官方岗位ID只从岗位链接里取：查询参数里的岗位ID，或最后一个含数字的路径片段；取不到即为 None。"""
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

    @staticmethod
    def origin_of(url):
        p = urlparse(str(url or ''))
        return f'{p.scheme}://{p.netloc}' if p.scheme in ('http', 'https') and p.netloc else None

    def company_id(self, sheet, name, used):
        latin = re.sub(r'[^a-z0-9]+', '-', str(name).lower()).strip('-')
        cid = latin if len(latin) >= 2 else 'c-' + sha(f'{sheet}:{name}'.encode())[:8]
        if sheet != '中国大陆':
            cid += '-' + {'香港': 'hk', '新加坡': 'sg'}.get(sheet, 'x')
        base, n = cid[:70], 2
        while cid in used and used[cid] != (sheet, name):
            cid = f'{base}-{n}'
            n += 1
        used[cid] = (sheet, name)
        return cid

    def is_third_party(self, url):
        host = urlparse(str(url or '')).netloc.lower()
        return any(host == h or host.endswith('.' + h) for h in self.third_party)

    def folder(self, sheet, name):
        return self.base / safe(sheet) / safe(name)

    def link_problems(self, link, rid, dup):
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
        return {v.get('公司'): v.get('是否分BU可平行投递') for _, v in S.quota_records(master) if v.get('范围') == '公司总量'}

    @staticmethod
    def group_of(parallel, bu):
        bu = str(bu or '').strip()
        return bu if parallel == '是' and bu and bu != '未披露' else '公司'

    @staticmethod
    def order_problems(roles):
        """同一志愿分组内：只有一个岗位时可不填；多个岗位时每个都必须有用户给定的、不重复的正整数顺序
        （可以不从1起，例如第1志愿已投递，本批只剩第2、3志愿）。"""
        groups = collections.defaultdict(list)
        for r in roles:
            groups[r.get('志愿分组') or '公司'].append(r.get('顺序'))
        bad = {}
        for g, orders in groups.items():
            if len(orders) == 1 and orders[0] in (None, ''):
                continue
            if not all(isinstance(o, int) and not isinstance(o, bool) and o > 0 for o in orders) \
                    or len(set(orders)) != len(orders):
                bad[g] = orders
        return bad

    # ---------- refresh ----------
    def refresh(self):
        """按主表覆盖重建全部 公司.json：只保留主表没有的运行字段，旧文件里其它内容一律丢弃。"""
        master = self.ctx.load_master()
        rows = self.master_rows(master)
        used, written, removed = {}, 0, []
        existing = {p.parent: p for p in self.base.glob('*/*/公司.json')}
        for p in existing.values():
            try:
                d = json.loads(p.read_text(encoding='utf-8'))
                used[d['companyId']] = (d['地区'], d['公司'])
            except (OSError, ValueError, KeyError):
                pass
        keep_folders = set()
        parallel = self.bu_parallel(master)
        for (sheet, name), rs in sorted(rows.items(), key=lambda x: (x[0][0], str(x[0][1]))):
            folder = self.folder(sheet, name)
            cf = folder / '公司.json'
            role_rows = [r for r in rs if r['kind'] == 'role']
            assets = sorted(folder.glob('登录*.json')) + sorted(folder.glob('表单定义*.json'))
            has_assets = bool(assets)
            if not role_rows and not has_assets:
                continue
            old = json.loads(cf.read_text(encoding='utf-8')) if cf.exists() else {}
            old_roles = {kid(r.get('主表行身份')): r for r in old.get('roles', [])}
            old_by_link = {str(r.get('岗位链接') or '').rstrip('/'): r for r in old.get('roles', []) if r.get('岗位链接')}
            roles = []
            for r in role_rows:
                link = r['values'].get(S.LINK_COL)
                prev = old_roles.get(kid(r['key'])) or old_by_link.get(str(link or '').rstrip('/')) or {}
                item = {'主表行身份': r['key'], '岗位ID': self.role_id(link), 'BU': r['values'].get('所属BU') or '未披露',
                        '志愿分组': self.group_of(parallel.get(name), r['values'].get('所属BU'))}
                for f in KEEP_ROLE_FIELDS:
                    if prev.get(f) not in (None, '', []):
                        item[f] = prev[f]
                roles.append(item)
            login_origins = set()
            for f in assets:
                try:
                    login_origins.add(json.loads(f.read_text(encoding='utf-8'))['origin'])
                except (OSError, ValueError, KeyError):
                    pass
            origins = sorted({o for o in (self.origin_of(r['values'].get(S.LINK_COL)) for r in role_rows)
                              if o and not self.is_third_party(o)} | set(old.get('origins') or []) | login_origins)
            cid = old.get('companyId') or self.company_id(sheet, name, used)
            doc = {'schemaVersion': 2, 'companyId': cid, '公司': name, '地区': sheet, 'origins': origins,
                   '志愿分组依据': '投递名额“公司总量”行 是否分BU可平行投递=' + str(parallel.get(name) or '无名额规则'),
                   'roles': roles, 'updatedAt': now()}
            if old.get('志愿顺序确认') and any('顺序' in r for r in roles):
                doc['志愿顺序确认'] = old['志愿顺序确认']
            folder.mkdir(parents=True, exist_ok=True)
            S.save_json(cf, doc)
            keep_folders.add(folder)
            written += 1
        for folder in sorted({p.parent for p in self.base.glob('*/*/公司.json')} - keep_folders):
            shutil.rmtree(folder)
            removed.append(str(folder.relative_to(self.root)))
        return {'written': written, 'folders_removed': len(removed)}

    # ---------- set-order ----------
    def set_order(self, items, quote):
        """用户给定志愿顺序：[{sheet,company,link|title,顺序}]；同公司一次给全，按分组校验后覆盖写。"""
        if not quote or not quote.strip():
            raise SystemExit('必须提供用户给定顺序的原话')
        by_company = collections.defaultdict(list)
        for it in items:
            if not isinstance(it.get('顺序'), int) or isinstance(it.get('顺序'), bool) or it['顺序'] < 1:
                raise SystemExit(f"顺序必须是正整数：{it}")
            by_company[(it['sheet'], it['company'])].append(it)
        out = []
        for (sheet, name), its in by_company.items():
            cf = self.folder(sheet, name) / '公司.json'
            if not cf.exists():
                raise SystemExit(f'缺少 {sheet}/{name} 的 公司.json，请先运行 refresh')
            doc = json.loads(cf.read_text(encoding='utf-8'))
            roles = {kid(r['主表行身份']): r for r in doc['roles']}
            touched = set()
            for it in its:
                key = S.normalize_key({k: it.get(k) for k in ('sheet', 'company', 'link', 'title')})
                hit = [r for r in doc['roles'] if all(r['主表行身份'].get(k) == v for k, v in key.items())]
                if len(hit) != 1:
                    raise SystemExit(f'找不到唯一岗位：{key}')
                hit[0]['顺序'] = it['顺序']
                touched.add(hit[0]['志愿分组'])
            # 本次涉及的分组，未给顺序的岗位清掉旧顺序，避免新旧顺序混用
            for r in doc['roles']:
                if r['志愿分组'] in touched and not any(
                        all(r['主表行身份'].get(k) == v for k, v in S.normalize_key({k: it.get(k) for k in ('sheet', 'company', 'link', 'title')}).items())
                        for it in its):
                    r.pop('顺序', None)
            bad = self.order_problems([r for r in doc['roles'] if r['志愿分组'] in touched and '顺序' in r])
            if bad:
                raise SystemExit(f'{name} 志愿顺序不合格（分组内须为不重复的正整数）：{bad}')
            doc['志愿顺序确认'] = {'userQuote': quote.strip(), 'confirmedAt': now()}
            doc['updatedAt'] = now()
            S.save_json(cf, doc)
            out.append({'公司': name, '分组': {g: sorted(((r['顺序'], r['主表行身份'].get('title')) for r in doc['roles']
                                                    if r['志愿分组'] == g and '顺序' in r)) for g in sorted(touched)}})
        return out

    # ---------- forward ----------
    def batch(self):
        if not self.batch_file.exists():
            return None
        return json.loads(self.batch_file.read_text(encoding='utf-8'))

    def forward(self, session=None, turn=None):
        from application_requirements import Requirements
        return Requirements(self.root, session, turn).build(self._forward)

    def _forward(self, requirements):
        from application_requirements import select_companies
        prior = self.batch()
        if prior and prior.get('state') not in STAGES_OK_FOR_NEW_BATCH:
            raise SystemExit(f"当前批次状态为「{prior.get('state')}」，只有 空闲/已回写 时才能生成新批次")
        master = self.ctx.load_master()
        rows = self.master_rows(master)
        companies, excluded = [], []
        ledger = self.ctx.ledger()
        for sheet, name in select_companies(rows, ledger, requirements):
            rs = rows[(sheet, name)]
            pending = [r for r in rs if r['kind'] == 'role' and r['values'].get(self.ctx.column) == '岗位待投递']
            if not pending:
                continue
            folder = self.folder(sheet, name)
            cf = folder / '公司.json'
            if not cf.exists():
                excluded.append({'公司': name, '地区': sheet, '原因': '缺少 公司.json，请先运行 refresh'})
                continue
            doc = json.loads(cf.read_text(encoding='utf-8'))
            by_key = {kid(r['主表行身份']): r for r in doc.get('roles', [])}
            links = collections.Counter(str(r['values'].get(S.LINK_COL) or '').rstrip('/') for r in rs if r['kind'] == 'role')
            ok = []
            for r in pending:
                link = r['values'].get(S.LINK_COL)
                rid = self.role_id(link)
                problems = self.link_problems(link, rid, links[str(link or '').rstrip('/')] > 1)
                note, confirmed = S.link_exception_note(self.ctx, ledger, sheet, r['row'], r['idx'])
                if note and not confirmed:
                    problems.append('备注是链接例外，但没有用户确认记录，须用 link-exception 登记')
                elif note:  # 用户确认的链接例外：官网无独立详情页，只免“列表页”检查
                    problems = [x for x in problems if x != LIST_PAGE_PROBLEM]
                if problems:
                    excluded.append({'公司': name, '地区': sheet, '岗位': r['values'].get(S.NAME_COL), '原因': '；'.join(problems)})
                    continue
                stored = by_key.get(kid(r['key']), {})
                item = {'岗位ID': rid, '岗位名称': r['values'].get(S.NAME_COL), '岗位链接': link,
                        'BU': stored.get('BU') or r['values'].get('所属BU') or '未披露', '志愿分组': stored.get('志愿分组') or '公司'}
                item['materials'] = role_materials(r['values'], master['data_sha256'])
                if note:
                    item['链接说明'] = note[len(S.LINK_EXCEPTION_PREFIX):]
                if stored.get('顺序') not in (None, ''):
                    item['顺序'] = stored['顺序']
                if stored.get('投递链接'):  # 投递组件进入申请页后写回的申请页链接：下批直接从申请页进入
                    item['投递链接'] = stored['投递链接']
                entry_problem = entry_link_problem(folder, item)
                if entry_problem:
                    excluded.append({'公司': name, '地区': sheet, '岗位': item['岗位名称'], '原因': entry_problem})
                    continue
                ok.append(item)
            bad = self.order_problems(ok)
            for g, orders in bad.items():
                for item in [i for i in ok if i['志愿分组'] == g]:
                    excluded.append({'公司': name, '地区': sheet, '岗位': item['岗位名称'],
                                     '原因': f'志愿分组「{g}」有 {len(orders)} 个待投递岗位，志愿顺序未由用户给全（现为 {orders}）；用 set-order 记录用户给定的顺序后重新生成'})
            ok = sorted([i for i in ok if i['志愿分组'] not in bad], key=lambda i: (i['志愿分组'], i.get('顺序') or 1))
            if ok:
                companies.append({'companyId': doc['companyId'], '公司': name, '地区': sheet,
                                  '公司文件': str(cf.relative_to(self.root)), '岗位': ok,
                                  'stage': '已建档' if any(folder.glob('登录*.json')) else '待建档'})
        batch = {'schemaVersion': 2, 'materialContractVersion': 1, 'batchId': 'B' + datetime.now().strftime('%Y%m%d%H%M%S'),
                 'state': '等待正向评审', 'updatedAt': now(), 'masterDataSha256': master['data_sha256'],
                 'companies': companies, 'excluded': excluded, 'confirmation': None}
        if requirements.get('浏览器'):
            batch['execution'] = {'browserChannels': ['chrome', 'iab'] if requirements['浏览器'] == 'dual'
                                  else [requirements['浏览器']]}
        S.save_json(self.batch_file, batch)
        self.review(batch)
        return {'batchId': batch['batchId'], 'companies': len(companies),
                'roles': sum(len(c['岗位']) for c in companies), 'excluded': len(excluded),
                'review': str(self.review_file.relative_to(self.root))}

    def amend_materials(self, quote=None, expected_sha=None):
        """Material-only amendment. Never reselect, reorder, or recreate queue roles."""
        b = self.batch()
        if not b or b.get('state') not in ('已交接', '进行中') or verify_handoff(b):
            raise SystemExit('必须是已确认且未变更的当前批次')
        prior_sha = content_sha(b)
        master = self.ctx.load_master()
        rows = self.master_rows(master)
        changes = []
        for c in b['companies']:
            for role in c['岗位']:
                hits = [r for r in rows.get((c['地区'], c['公司']), []) if r['kind'] == 'role'
                        and r['values'].get(S.LINK_COL) == role['岗位链接']
                        and r['values'].get(S.NAME_COL) == role['岗位名称']]
                if len(hits) != 1:
                    raise SystemExit('当前岗位无法唯一对应主表，拒绝补充材料')
                materials = role_materials(hits[0]['values'], master['data_sha256'])
                if materials != role.get('materials'):
                    changes.append({'公司': c['公司'], '岗位': role['岗位名称'],
                                    'current': role.get('materials'), 'updated': materials})
                    role['materials'] = materials
        b['materialContractVersion'] = 1
        preview = {'batchId': b['batchId'], 'previousContentSha256': prior_sha, 'changes': changes,
                   'newContentSha256': content_sha(b), 'selectionChanged': False}
        if quote is None:
            return preview
        if not quote.strip() or expected_sha != prior_sha:
            raise SystemExit('必须附用户授权原话和预览时的原批次摘要')
        b['updatedAt'] = now()
        set_confirmation(b, 'materials', quote.strip(), b['updatedAt'], previousContentSha256=prior_sha)
        S.save_json(self.batch_file, b)
        self.review(b)
        if verify_handoff(self.batch()):
            raise ValueError('补充材料后批次确认读回失败')
        return preview

    def review(self, batch):
        rows = ''.join(
            f"<tr><td>{html.escape(c['地区'])}</td><td>{html.escape(c['公司'])}</td><td>{html.escape(c['stage'])}</td>"
            f"<td>{'<br>'.join(html.escape(('' if r['志愿分组'] == '公司' else r['志愿分组'] + '｜') + ('第' + str(r['顺序']) + '志愿｜' if r.get('顺序') else '') + str(r['岗位名称']) + ('（链接例外：' + str(r['链接说明']) + '）' if r.get('链接说明') else '') + '｜' + material_summary(r)) for r in c['岗位'])}</td></tr>" for c in batch['companies'])
        ex = ''.join(f"<tr><td>{html.escape(e['地区'])}</td><td>{html.escape(str(e['公司']))}</td><td>{html.escape(str(e.get('岗位', '')))}</td>"
                     f"<td>{html.escape(e['原因'])}</td></tr>" for e in batch['excluded'])
        page = f"""<!doctype html><html lang=zh><meta charset=utf-8><meta name=viewport content="width=device-width,initial-scale=1"><title>正向评审</title>
<style>body{{font:14px/1.5 -apple-system,"PingFang SC",sans-serif;margin:16px;max-width:1100px}}table{{border-collapse:collapse;width:100%;font-size:13px;margin:8px 0 20px}}td,th{{border-top:1px solid #e5e5e5;padding:5px 6px;text-align:left;vertical-align:top}}th{{background:#f5f5f5}}</style>
<h1>正向评审 · 批次 {html.escape(batch['batchId'])}</h1>
<p>状态：{html.escape(batch['state'])} · 生成时间：{html.escape(batch['updatedAt'])} · 可交接 {len(batch['companies'])} 家、{sum(len(c['岗位']) for c in batch['companies'])} 个岗位；不交接 {len(batch['excluded'])} 项。</p>
<h2>交接给投递组件</h2><table><tr><th>地区</th><th>公司</th><th>建档</th><th>岗位</th></tr>{rows}</table>
<h2>本批不交接（需先处理）</h2><table><tr><th>地区</th><th>公司</th><th>岗位</th><th>原因</th></tr>{ex}</table></html>"""
        self.review_file.parent.mkdir(parents=True, exist_ok=True)
        self.review_file.write_text(page, encoding='utf-8')
        if self.review_file.read_text(encoding='utf-8') != page:
            raise ValueError('正向评审写入后读回不一致')

    def reverse(self):
        """反向评审：读取投递回传，逐行给出“当前状态是X，更新状态是Y”，生成反向评审页；不改主表。"""
        b = self.batch()
        if not b or b.get('state') not in ('已交接', '进行中', '等待反向评审'):
            raise SystemExit('没有已交接的投递批次')
        items, notes, _ = _reverse_items(self)
        master = self.ctx.load_master()
        plan = S.propose(self.ctx, [{k: v for k, v in i.items() if k != 'outcomeKey'} for i in items], master) if items else []
        lines = [p['text'] for p in plan]
        b['state'], b['updatedAt'] = '等待反向评审', now()
        S.save_json(self.batch_file, b)
        page = f"""<!doctype html><html lang=zh><meta charset=utf-8><meta name=viewport content="width=device-width,initial-scale=1"><title>反向评审</title>
<style>body{{font:14px/1.5 -apple-system,"PingFang SC",sans-serif;margin:16px;max-width:1100px}}li{{margin:4px 0}}</style>
<h1>反向评审 · 批次 {html.escape(b['batchId'])}</h1><p>生成时间：{html.escape(b['updatedAt'])}</p>
<h2>主表状态变更（需你确认）</h2><ul>{''.join('<li>' + html.escape(t) + '</li>' for t in lines) or '<li>无</li>'}</ul>
<h2>不改状态、需要你知道或处理的结果</h2><ul>{''.join('<li>' + html.escape(n['title'] + '：' + n['说明']) + '</li>' for n in notes) or '<li>无</li>'}</ul></html>"""
        self.batch_file.with_name('反向评审.html').write_text(page, encoding='utf-8')
        return {'batchId': b['batchId'], 'state': b['state'], 'show_user_exactly': lines,
                'notes': [n['title'] + '：' + n['说明'] for n in notes], 'review': '岗位交接/反向评审.html'}

    def relink_observed(self, quote=None, expected_sha=None):
        """投递端发现岗位链接已变（网站跳转或改链接）时的唯一更新路径：预览“岗位链接 X → Y”，
        用户确认后用 relink 覆盖主表链接（旧链接只留 git 历史），refresh 按新链接重算岗位ID；
        岗位在当前批次时同步改批次（链接作用域确认），执行端须重新导入队列。"""
        f = _outcome_file(self)
        out = json.loads(f.read_text(encoding='utf-8')) if f.exists() else {'schemaVersion': 1, 'roles': {}}
        proposals = out.get('linkChanges') or {}
        if not proposals:
            return {'show_user_exactly': [], 'note': '没有待确认的链接变更'}
        docs = {}
        for cf in self.base.glob('*/*/公司.json'):
            d = json.loads(cf.read_text(encoding='utf-8'))
            docs[d['companyId']] = d
        b = self.batch()
        active = b if b and b.get('state') in ('已交接', '进行中', '等待反向评审') else None
        items, lines, batch_edits = [], [], []
        for key, p in sorted(proposals.items()):
            d = docs.get(p['companyId'])
            hit = [r['主表行身份'] for r in (d or {}).get('roles', [])
                   if str((r.get('主表行身份') or {}).get('link') or '').rstrip('/') == str(p['roleUrl']).rstrip('/')]
            if len(hit) != 1:
                raise SystemExit(f"链接变更 {key} 对不上主表岗位（主表链接可能已改），请人工核对")
            items.append({**{k: v for k, v in hit[0].items() if k != 'placeholder'}, 'new_link': p['observedUrl']})
            if p.get('evidence') == 'derived_from_application_page':
                lines.append(f"{key}：新链接是从已打开的申请页推出来的，还没在详情页核实过；确认后下次从这个链接进入时会在详情页核对岗位")
            new_id = self.role_id(p['observedUrl'])
            if not new_id:
                raise SystemExit(f"新链接里没有官方岗位ID：{p['observedUrl']}")
            if active:
                for c in active['companies']:
                    for r in c['岗位']:
                        if c['companyId'] == p['companyId'] and str(r['岗位链接']).rstrip('/') == str(p['roleUrl']).rstrip('/'):
                            batch_edits.append((c, r, p['observedUrl'], new_id))
                            lines.append(f"批次 {active['batchId']}：{c['公司']}｜{r['岗位名称']} 岗位链接「{r['岗位链接']}」→「{p['observedUrl']}」，岗位ID「{r['岗位ID']}」→「{new_id}」（批次登记链接作用域确认，执行端须重新导入队列）")
        preview = S.relink(self.ctx, items, None, preview=True, allow_in_batch=True)
        prior_sha = content_sha(active) if active and batch_edits else None
        shown = preview['show_user_exactly'] + lines
        if quote is None:
            return {'show_user_exactly': shown, 'previousContentSha256': prior_sha,
                    'rule': '原样展示给用户并取得明确确认后，带 --quote 与 --expected-sha 执行'}
        if not quote.strip() or (prior_sha and expected_sha != prior_sha):
            raise SystemExit('必须附用户确认原话，以及预览时的批次摘要')
        S.relink(self.ctx, items, quote.strip(), allow_in_batch=True)
        self.refresh()
        renamed = {}
        if batch_edits:
            master = self.ctx.load_master()
            rows = self.master_rows(master)
            for c, r, new_link, new_id in batch_edits:
                old_key = c['companyId'] + '|' + str(r['岗位ID'])
                hits = [x for x in rows.get((c['地区'], c['公司']), []) if x['kind'] == 'role'
                        and x['values'].get(S.LINK_COL) == new_link and x['values'].get(S.NAME_COL) == r['岗位名称']]
                if len(hits) != 1:
                    raise SystemExit('换链接后主表岗位无法唯一对应，批次未改')
                r['岗位链接'], r['岗位ID'] = new_link, new_id
                r['materials'] = role_materials(hits[0]['values'], master['data_sha256'])
                renamed[old_key] = c['companyId'] + '|' + new_id
            prog = active.get('progress') or {}
            for k in ('processedOutcomeKeys', 'remainingRoleKeys'):
                if k in prog:
                    prog[k] = sorted(renamed.get(x, x) for x in prog[k])
            active['updatedAt'] = now()
            set_confirmation(active, 'links', quote.strip(), active['updatedAt'], previousContentSha256=prior_sha)
            S.save_json(self.batch_file, active)
            self.review(active)
            if verify_handoff(self.batch()):
                raise ValueError('换链接后批次确认读回失败')
        roles = {renamed.get(k, k): ({**v, 'roleId': renamed[k].split('|', 1)[1]} if k in renamed else v) for k, v in (out.get('roles') or {}).items()}
        rest = {'schemaVersion': 1, 'roles': roles}
        if roles:
            S.save_json(f, rest)
        elif f.exists():
            f.unlink()
        return {'relinked': len(items), 'batchUpdated': bool(batch_edits), 'show_user_exactly': shown}

    def reverse_confirm(self, quote, resolved=()):
        """用户确认后：写主表状态（经 json_excel_sync 确认台账），清除已处理的回传；全部处理完批次 → 已回写。"""
        if not quote or not quote.strip():
            raise SystemExit('必须提供用户确认原话')
        b = self.batch()
        if not b or b.get('state') != '等待反向评审':
            raise SystemExit('没有等待反向评审的批次，请先运行 reverse')
        if (b.get('confirmation') or {}).get('contentSha256') != content_sha(b):
            raise SystemExit('投递批次在确认后被改动，须重新正向评审')
        items, notes, roles = _reverse_items(self)
        if items:
            S.confirm(self.ctx, [{k: v for k, v in i.items() if k != 'outcomeKey'} for i in items], quote.strip(), f"反向评审 {b['batchId']}")
        done = {i['outcomeKey'] for i in items} | {n['key'] for n in notes if n['event'] not in UNRESOLVED} | set(resolved)
        left = {k: v for k, v in roles.items() if k not in done}
        f = _outcome_file(self)
        link_changes = (json.loads(f.read_text(encoding='utf-8')).get('linkChanges') or {}) if f.exists() else {}
        if left or link_changes:  # 链接变更提议另走 relink-observed，反向评审不得清掉
            S.save_json(f, {'schemaVersion': 1, 'roles': left, **({'linkChanges': link_changes} if link_changes else {})})
        elif f.exists():
            f.unlink()
        processed = set(b.get('progress', {}).get('processedOutcomeKeys', [])) | done
        rows = self.master_rows(self.ctx.load_master())
        pending = []
        for company in b['companies']:
            for role in company['岗位']:
                key = company['companyId'] + '|' + role['岗位ID']
                matches = [r for r in rows.get((company['地区'], company['公司']), [])
                           if r['values'].get(S.LINK_COL) == role['岗位链接']]
                closed = len(matches) == 1 and matches[0]['values'].get(self.ctx.column) not in ('岗位待投递', '岗位待筛选', None, '')
                if key not in processed and not closed:
                    pending.append(key)
        b['progress'] = {'processedOutcomeKeys': sorted(processed), 'remainingRoleKeys': sorted(pending)}
        b['state'] = '等待反向评审' if left else '进行中' if pending else '已回写'
        b['updatedAt'] = now()
        S.save_json(self.batch_file, b)
        if b['state'] == '已回写':
            from application_requirements import Requirements
            Requirements(self.root).complete(b['batchId'])
        return {'batchId': b['batchId'], 'state': b['state'], 'statusChanges': len(items), 'unresolved': sorted(left), 'remainingRoles': len(pending)}

    def forward_confirm(self, quote):
        if not quote or not quote.strip():
            raise SystemExit('必须提供用户确认原话')
        b = self.batch()
        if not b or b.get('state') != '等待正向评审':
            raise SystemExit('没有等待正向评审的批次')
        if b.get('materialContractVersion') != 1:
            raise SystemExit('投递材料契约缺失，请重新生成正向评审')
        master = self.ctx.load_master()
        rows = self.master_rows(master)
        for company in b['companies']:
            for role in company['岗位']:
                hits = [r for r in rows.get((company['地区'], company['公司']), []) if r['kind'] == 'role'
                        and r['values'].get(S.LINK_COL) == role['岗位链接'] and r['values'].get(S.NAME_COL) == role['岗位名称']]
                if len(hits) != 1 or role.get('materials') != role_materials(hits[0]['values'], master['data_sha256']):
                    raise SystemExit('岗位材料遗漏或来源发生变化，请重新生成正向评审')
        at = now()
        b['confirmation'] = None
        set_confirmation(b, 'selection', quote.strip(), at)
        b['state'], b['updatedAt'] = '已交接', at
        S.save_json(self.batch_file, b)
        return {'batchId': b['batchId'], 'state': b['state']}


REVERSE_STATUS = {'submission_confirmed': '岗位已投递', 'role_closed': '岗位失效/关闭'}
UNRESOLVED = ('submission_uncertain', 'application_blocked')


def _outcome_file(st):
    return st.batch_file.with_name('投递回传.json')


def _reverse_items(st):
    """把 投递回传.json 的每条结果对到主表行：已投递/已关闭 → 状态变更；不确定/受阻/跳过 → 只列出待用户处理。"""
    f = _outcome_file(st)
    roles = json.loads(f.read_text(encoding='utf-8')).get('roles', {}) if f.exists() else {}
    docs = {}
    for cf in st.base.glob('*/*/公司.json'):
        d = json.loads(cf.read_text(encoding='utf-8'))
        docs[d['companyId']] = d
    master = st.ctx.load_master()
    rows = st.master_rows(master)
    items, notes = [], []
    for key, o in sorted(roles.items()):
        d = docs.get(o['companyId'])
        k = None
        if d:
            hit = [r['主表行身份'] for r in d.get('roles', []) if r.get('岗位ID') == o['roleId']]
            if len(hit) == 1:
                k = hit[0]
            else:
                cand = [r['key'] for r in rows.get((d['地区'], d['公司']), []) if r['values'].get(S.LINK_COL) == o['roleUrl']]
                k = cand[0] if len(cand) == 1 else None
        if not k:
            notes.append({'key': key, 'event': o['event'], 'title': o['roleTitle'], '说明': '回传结果对不上主表岗位，需人工核对'})
            continue
        if o['event'] in REVERSE_STATUS:
            items.append({**{x: v for x, v in k.items() if x != 'placeholder'}, 'to': REVERSE_STATUS[o['event']], 'outcomeKey': key})
        else:
            notes.append({'key': key, 'event': o['event'], 'title': o['roleTitle'],
                          '说明': '投递结果不确定或受阻，处理前禁止重跑' if o['event'] in UNRESOLVED else '用户跳过，状态不变'})
    return items, notes, roles


# 确认按作用域保存：选岗（selection）与补材料（materials）各留最新一条，互不覆盖。
# 顶层 userQuote/confirmedAt/contentSha256 是当前生效的门禁（最近一次确认），供投递端核对；
# 顶层 scope 沿用投递端约定：materials 作用域写 'materials_only'。
SCOPE_LABEL = {'selection': 'selection', 'materials': 'materials_only', 'links': 'links_only'}


def set_confirmation(b, scope, quote, at, **extra):
    prior = b.get('confirmation') or {}
    scopes = dict(prior.get('scopes') or {})
    if prior and not scopes and prior.get('userQuote') and prior.get('scope') != 'materials_only':
        scopes['selection'] = {'userQuote': prior['userQuote'], 'confirmedAt': prior.get('confirmedAt'),
                               'contentSha256': prior.get('contentSha256')}
    current = content_sha(b)
    scopes[scope] = {'userQuote': quote, 'confirmedAt': at, 'contentSha256': current, **extra}
    b['confirmation'] = {'userQuote': quote, 'confirmedAt': at, 'contentSha256': current,
                         'scope': SCOPE_LABEL[scope], **extra, 'scopes': scopes}


def content_sha(batch):
    body = {k: v for k, v in batch.items() if k not in ('confirmation', 'state', 'updatedAt', 'progress')}
    return sha(json.dumps(body, ensure_ascii=False, sort_keys=True).encode())


def verify_handoff(batch):
    """供 Hook/投递组件调用：批次已交接且内容未被改动。"""
    if not batch or batch.get('state') not in ('已交接', '进行中'):
        return '投递批次未经用户正向评审确认'
    if (batch.get('confirmation') or {}).get('contentSha256') != content_sha(batch):
        return '投递批次在确认后被改动，须重新正向评审'
    return None


def main():
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--workspace', default=os.environ.get('CAMPUS_RECRUITING_WORKSPACE', '.'))
    ap.add_argument('--labels', default=os.environ.get('CAMPUS_STATUS_LABELS', '产品管理/投递状态标签.json'))
    ap.add_argument('--session', help='当前宿主会话身份，默认使用宿主环境')
    ap.add_argument('--turn', help='当前宿主回合身份，默认读取需求Hook上下文')
    sub = ap.add_subparsers(dest='cmd', required=True)
    sub.add_parser('refresh')
    sub.add_parser('forward')
    p = sub.add_parser('amend-materials')
    p.add_argument('--quote', help='不传则仅预览；传用户明确授权原话后只补当前岗位材料')
    p.add_argument('--expected-sha', help='预览返回的 previousContentSha256')
    p = sub.add_parser('relink-observed')
    p.add_argument('--quote', help='不传则仅预览；传用户确认原话后覆盖主表链接并同步批次')
    p.add_argument('--expected-sha', help='预览返回的 previousContentSha256（岗位在当前批次时必填）')
    p = sub.add_parser('set-order')
    p.add_argument('--input', required=True, help='UTF-8 JSON：[{sheet,company,link|title,顺序}]，同公司同分组一次给全')
    p.add_argument('--quote', required=True, help='用户给定顺序的原话')
    p = sub.add_parser('forward-confirm')
    p.add_argument('--quote', required=True)
    sub.add_parser('batch-status')
    sub.add_parser('reverse')
    p = sub.add_parser('reverse-confirm')
    p.add_argument('--quote', required=True)
    p.add_argument('--resolved', action='append', default=[], help='用户已处理的不确定/受阻结果（companyId|岗位ID），确认后清除')
    a = ap.parse_args()
    st = Store(a.workspace, a.labels)
    if a.cmd == 'refresh':
        with S.exclusive(st.ctx):
            out = st.refresh()
    elif a.cmd == 'forward':
        with S.exclusive(st.ctx):
            out = st.forward(a.session, a.turn)
    elif a.cmd == 'amend-materials':
        with S.exclusive(st.ctx):
            out = st.amend_materials(a.quote, a.expected_sha)
    elif a.cmd == 'relink-observed':
        with S.exclusive(st.ctx):
            out = st.relink_observed(a.quote, a.expected_sha)
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
        b = st.batch()
        out = {'batchId': (b or {}).get('batchId'), 'state': (b or {}).get('state'), 'handoffProblem': verify_handoff(b)}
    print(json.dumps(out, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
