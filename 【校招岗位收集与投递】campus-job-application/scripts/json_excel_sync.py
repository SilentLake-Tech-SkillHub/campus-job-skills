#!/usr/bin/env python3
"""JSON→Excel 唯一同步脚本：工作簿全量 JSON 为唯一主数据，Excel 只由本脚本从 JSON 写出。

投递状态只能使用标签配置 JSON 中定义的标签；任何状态变更都必须先有用户对
“当前状态是X，更新状态是Y”的逐行确认记录（确认台账），否则拒绝写入。

子命令：
  check     校验主数据标签规则，报告 JSON 与 Excel 的差异和未确认的状态变更
  status    只读：显示某行当前状态
  propose   只读：生成“当前状态是X，更新状态是Y”的核对文本，不改任何文件
  confirm   用户明确确认后，把变更逐行记入确认台账并更新主数据 JSON（不写 Excel）
  report    只读：一家或多家公司的汇报（主链接、计划、管培/产培、岗位按状态、名额与占用）
  upsert    按业务键增改主数据 JSON 的非状态字段（新岗位、JD 等）；不能改投递状态，也不能更换已有岗位的链接
  link-exception  登记/撤销“官网无独立详情页”的链接例外：备注只能由本命令写成固定格式，用户确认后执行；forward 据此不检查该岗位链接是否为具体岗位页
  relink    更换已有岗位的链接：--preview 生成“岗位链接 X → Y”给用户看，确认后带 --quote 执行；状态不变、无需重新确认，志愿顺序等投递数据随行迁移
  write     JSON→Excel：标签门禁→状态确认门禁→写出→下拉校验→逐格读回→回写绑定（不做快照，主数据只留一份“上一版”）
  defer     用户选择“下一轮再写入”时登记延后（绑定当前 JSON 数据摘要）
  pending   供 Hook 使用：是否存在未写入 Excel 的 JSON 改动、用户是否已选择延后
"""
import argparse
import re
import contextlib
import fcntl
import copy
import hashlib
import io
import json
import os
import sys
import uuid
from datetime import date, datetime, time, timezone
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

NAME_COL, LINK_COL, COMPANY_COL = '岗位名称', '岗位链接', '公司'
COMPANY_INFO_COLS = ('地区', COMPANY_COL, '主要行业', '企业性质')
NOTE_COL = '备注'
# 链接例外：官网没有独立的岗位详情页时，备注写成这个固定开头 + 定位说明，forward 不再检查该岗位链接是否为具体岗位页。
# 这种备注只能由 link-exception 命令写入（用户确认、记入台账），upsert 不能写或改。
LINK_EXCEPTION_PREFIX = '用户例外保留：官网无独立详情页。'
ROW_KINDS = {'role': '岗位行', 'company': '公司行'}


def load_label_groups(root, cfg, cfg_path):
    """标签统一加载入口（所有脚本只经这里读标签）：主配置本身是岗位组，labelGroups 列出其它组文件（一组一个文件）。
    每组只适用于一种行类型（rowKind）；各组标签不得重名。返回 (标签→适用行类型集合, 各组列表)。"""
    root = Path(root)
    groups = []
    for path, g in [(Path(cfg_path), cfg)] + [(root / p, None) for p in cfg.get('labelGroups', [])]:
        if g is None:
            try:
                g = json.loads(path.read_text(encoding='utf-8'))
            except (OSError, ValueError) as e:
                raise SystemExit(f'标签组文件不可读：{path}（{e}）')
        if g.get('schemaVersion') != 1:
            raise SystemExit(f'标签组版本不支持：{path}')
        if g.get('rowKind') not in ROW_KINDS:
            raise SystemExit(f'标签组 rowKind 只能是 {"/".join(ROW_KINDS)}：{path}')
        labels = g.get('labels') or []
        if not labels or any(not l.get('value') for l in labels):
            raise SystemExit(f'标签组为空或含空标签：{path}')
        groups.append({'path': str(path), 'name': g.get('name', ''), 'rowKind': g['rowKind'], 'labels': labels})
    out = {}
    for grp in groups:
        for l in grp['labels']:
            if l['value'] in out:
                raise SystemExit(f'标签「{l["value"]}」在多个标签组重复定义')
            out[l['value']] = {grp['rowKind']}
    return out, groups


def label_problem(ctx, value, kind):
    """标签写入前检查：所有写入或校验投递状态的命令都经过这里。返回错误描述或 None。"""
    name = ROW_KINDS.get(kind, kind)
    if value in (None, ''):
        return '投递状态为空'
    if value not in ctx.labels:
        return f'非法标签「{value}」（不在任何标签组中）'
    if kind not in ctx.labels[value]:
        return f'标签「{value}」不适用于{name}；{name}只能用：' + '、'.join(ctx.kind_labels(kind))
    return None


def now():
    return datetime.now(timezone.utc).isoformat(timespec='seconds')


def digest(obj):
    return hashlib.sha256(json.dumps(obj, ensure_ascii=False, sort_keys=True, separators=(',', ':')).encode()).hexdigest()


def file_sha(p):
    return hashlib.sha256(Path(p).read_bytes()).hexdigest()


def write_bytes(path, data):
    """同目录临时文件后替换；挂载盘禁止替换时退回原地覆盖（Excel 可由主数据重写恢复）。"""
    path = Path(path)
    tmp = path.with_name('.' + path.name + '.' + uuid.uuid4().hex + '.tmp')
    try:
        tmp.write_bytes(data)
        os.replace(tmp, path)
    except PermissionError:
        with open(path, 'r+b' if path.exists() else 'wb') as f:
            f.write(data)
            f.truncate()
        try:
            tmp.unlink()
        except OSError:
            pass


def save_master(ctx, obj):
    """主数据只有一份，覆盖写；不另存任何“上一版”或快照，回滚靠 git。"""
    refresh_quota_counts(ctx, obj)
    refresh_digest(obj)
    save_json(ctx.master, obj)


def save_json(path, obj):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    write_bytes(path, json.dumps(obj, ensure_ascii=False, indent=2).encode('utf-8'))
    if json.loads(Path(path).read_text(encoding='utf-8')) != obj:
        raise SystemExit('JSON 读回不一致: ' + str(path))


class Ctx:
    def __init__(self, workspace, labels):
        self.root = Path(workspace).resolve()
        self.labels_path = (self.root / labels).resolve()
        self.cfg = json.loads(self.labels_path.read_text(encoding='utf-8'))
        if self.cfg.get('schemaVersion') != 1:
            raise SystemExit('标签配置版本不支持')
        self.column = self.cfg['column']
        self.sheets = self.cfg['sheets']
        self.labels, self.label_groups = load_label_groups(self.root, self.cfg, self.labels_path)
        self.label_files = [g['path'] for g in self.label_groups]
        self.master = self.root / self.cfg['masterData']
        self.book = self.root / self.cfg['workbook']
        self.ledger_path = self.root / self.cfg['confirmationLedger']
        fr = self.cfg.get('fieldRules')
        self.field_rules = json.loads((self.root / fr).read_text(encoding='utf-8'))['fields'] if fr and (self.root / fr).exists() else {}

    def kind_labels(self, kind):
        return [v for v, kinds in self.labels.items() if kind in kinds]

    def load_master(self):
        return json.loads(self.master.read_text(encoding='utf-8'))

    def ledger(self):
        if not self.ledger_path.exists():
            return {'schemaVersion': 1, 'entries': [], 'deferrals': []}
        return json.loads(self.ledger_path.read_text(encoding='utf-8'))


@contextlib.contextmanager
def exclusive(ctx, timeout=60):
    """多个 Agent 并发时串行化“读-改-写”：同一工作区内同一时刻只允许一个写命令。"""
    import time as _t
    lock = ctx.labels_path.parent / '.json_excel_sync.lock'
    with open(lock, 'a+') as f:
        deadline = _t.time() + timeout
        while True:
            try:
                fcntl.flock(f, fcntl.LOCK_EX | fcntl.LOCK_NB)
                break
            except BlockingIOError:
                if _t.time() > deadline:
                    raise SystemExit('另一个 Agent 正在写主数据，等待超时，请稍后重试')
                _t.sleep(0.2)
        try:
            yield
        finally:
            fcntl.flock(f, fcntl.LOCK_UN)


# ---------- 主数据 JSON 的行视图 ----------

def sheet_of(master, name):
    for s in master['sheets']:
        if s['name'] == name:
            return s
    raise SystemExit('主数据缺少工作表: ' + name)


def headers(sheet):
    first = next((r for r in sheet['rows'] if r['row'] == 1), None)
    if not first:
        raise SystemExit('工作表缺少表头: ' + sheet['name'])
    names = [c['value'] for c in first['cells']]
    if len(names) != len(set(names)):
        raise SystemExit('表头重复: ' + sheet['name'])
    return {c['value']: c['column'] for c in first['cells']}


def cell_value(row, col):
    for c in row['cells']:
        if c['column'] == col:
            return c.get('value')
    return None


def set_cell(row, col, value):
    for c in row['cells']:
        if c['column'] == col:
            if value in (None, ''):
                if c.get('hyperlink') or c.get('note'):
                    c['value'] = None
                else:
                    row['cells'].remove(c)
            else:
                c['value'] = value
                c['type'] = 's' if isinstance(value, str) else c.get('type', 'n')
            return
    if value not in (None, ''):
        row['cells'].append({'column': col, 'value': value, 'type': 's' if isinstance(value, str) else 'n'})
        row['cells'].sort(key=lambda c: c['column'])


def row_kind(row, idx):
    name = cell_value(row, idx[NAME_COL]) if NAME_COL in idx else None
    link = cell_value(row, idx[LINK_COL]) if LINK_COL in idx else None
    return 'role' if (name or link) else 'company'


def row_key(sheet_name, row, idx):
    company = cell_value(row, idx[COMPANY_COL])
    link = cell_value(row, idx[LINK_COL]) if LINK_COL in idx else None
    name = cell_value(row, idx[NAME_COL]) if NAME_COL in idx else None
    k = {'sheet': sheet_name, 'company': company}
    if link:
        k['link'] = link
    if name:
        k['title'] = name
    if not link and not name:
        k['placeholder'] = True
    return k


def key_id(k):
    return json.dumps(k, ensure_ascii=False, sort_keys=True)


def has_content(cells):
    return any(c.get('value') not in (None, '') for c in cells)


def iter_status_rows(ctx, master):
    for name in ctx.sheets:
        s = sheet_of(master, name)
        idx = headers(s)
        if ctx.column not in idx or COMPANY_COL not in idx:
            raise SystemExit(f'{name} 缺少 {ctx.column}/{COMPANY_COL} 列')
        for row in s['rows']:
            if row['row'] == 1 or not has_content(row['cells']):
                continue
            yield name, s, idx, row


def validate_labels(ctx, master):
    """标签门禁：每行状态必须是所属行类型的标签组中的标签；每家公司在每张地区表有且只有一行公司行。"""
    problems, seen, company_rows = [], {}, {}
    for name, s, idx, row in iter_status_rows(ctx, master):
        k = row_key(name, row, idx)
        kid = key_id(k)
        if not k['company']:
            problems.append(f'{name}!{row["row"]} 公司为空')
            continue
        if kid in seen:
            problems.append(f'{name}!{row["row"]} 与第{seen[kid]}行业务键重复：{kid}')
        seen[kid] = row['row']
        kind = row_kind(row, idx)
        company_rows[(name, k['company'])] = company_rows.get((name, k['company']), 0) + (kind == 'company')
        err = label_problem(ctx, cell_value(row, idx[ctx.column]), kind)
        if err:
            problems.append(f'{name}!{row["row"]} {k["company"]} {err}')
    for (name, company), n in company_rows.items():
        if not n:
            problems.append(f'{name}｜{company} 缺少公司行（每家公司在每张地区表固定一行公司行，有没有岗位都保留）')
    return problems


def link_exception_note(ctx, ledger, sheet_name, row, idx):
    """该岗位行是否为已确认的链接例外：备注是固定开头，且台账里有用户对这条备注的确认记录。返回 (备注, 是否已确认)。"""
    note = str(cell_value(row, idx[NOTE_COL]) or '') if NOTE_COL in idx else ''
    if not note.startswith(LINK_EXCEPTION_PREFIX):
        return None, False
    k = row_key(sheet_name, row, idx)
    for e in ledger.get('link_exceptions', []):
        ek = e.get('key') or {}
        same = ek.get('sheet') == k['sheet'] and ek.get('company') == k['company'] and (
            ek.get('link') == k['link'] if k.get('link') else ek.get('title') == k.get('title'))
        if same and e.get('note') == note:
            return note, True
    return note, False


def validate_link_exceptions(ctx, master):
    ledger, problems = ctx.ledger(), []
    for name, s, idx, row in iter_status_rows(ctx, master):
        if row_kind(row, idx) != 'role':
            continue
        note, ok = link_exception_note(ctx, ledger, name, row, idx)
        if note and not ok:
            problems.append(f'{name}!{row["row"]} {cell_value(row, idx[COMPANY_COL])} 备注是链接例外，但没有用户确认记录（或确认后被改动），须用 link-exception 登记')
    return problems


def check_field(ctx, field, value, kind):
    """按 表格字段规则.json 校验单个字段；返回错误描述或 None。"""
    rule = ctx.field_rules.get(field)
    if not rule or kind not in rule.get('适用行', list(ROW_KINDS)):
        if rule and value not in (None, ''):
            return f'「{field}」不适用于该行类型，应为空'
        return None
    if value in (None, ''):
        return f'「{field}」岗位行必填' if rule.get('岗位行必填') else None
    v = str(value)
    if rule.get('最大长度') and len(v) > rule['最大长度']:
        return f'「{field}」超过 {rule["最大长度"]} 字'
    if rule.get('格式') and not re.match(rule['格式'], v, re.S):
        return f'「{field}」格式不符：{rule.get("用途", "")}'
    for ban in rule.get('禁止模式', []):
        if re.search(ban['模式'], v):
            return f'「{field}」{ban["原因"]}：「{v[:60]}」'
    return None


def validate_fields(ctx, master):
    problems = []
    if not ctx.field_rules:
        return problems
    for name in ctx.sheets:
        s = sheet_of(master, name)
        idx = headers(s)
        for field in ctx.field_rules:
            if field not in idx:
                problems.append(f'{name} 缺少列「{field}」')
        for row in s['rows']:
            if row['row'] == 1 or not has_content(row['cells']):
                continue
            kind = row_kind(row, idx)
            for field in ctx.field_rules:
                if field in idx:
                    err = check_field(ctx, field, cell_value(row, idx[field]), kind)
                    if err:
                        problems.append(f'{name}!{row["row"]} {cell_value(row, idx[COMPANY_COL])} {err}')
    return problems


def add_column(ctx, sheet_name, column_name):
    master = ctx.load_master()
    s = sheet_of(master, sheet_name)
    idx = headers(s)
    if column_name in idx:
        return {'added': False, 'column': idx[column_name]}
    col = max(idx.values()) + 1
    first = next(r for r in s['rows'] if r['row'] == 1)
    template = next(c for c in first['cells'] if c['column'] == max(idx.values()))
    first['cells'].append({'column': col, 'value': column_name, 'type': 's'})
    s['max_column'] = max(s.get('max_column', 0), col)
    refresh_digest(master)
    save_master(ctx, master)
    return {'added': True, 'column': col, 'style_from_column': template['column']}


def find_row(ctx, master, key):
    s = sheet_of(master, key['sheet'])
    idx = headers(s)
    hits = [r for r in s['rows'] if r['row'] > 1 and key_id(row_key(key['sheet'], r, idx)) == key_id(key)]
    if len(hits) != 1:
        raise SystemExit(f'业务键匹配 {len(hits)} 行，需唯一：{key_id(key)}')
    return s, idx, hits[0]


def normalize_key(d):
    k = {'sheet': d['sheet'], 'company': d['company']}
    if d.get('link'):
        k['link'] = d['link']
    if d.get('title'):
        k['title'] = d['title']
    if not d.get('link') and not d.get('title'):
        k['placeholder'] = True
    return k


def refresh_digest(master):
    master['data_sha256'] = digest({'sheets': master['sheets'], 'company_batches': master.get('company_batches', {})})


def describe(key):
    return f"{key['sheet']}｜{key['company']}｜{key.get('title') or key.get('link') or '公司行'}"


# ---------- Excel 读写 ----------

def to_excel_value(v):
    if isinstance(v, dict) and set(v) == {'type', 'value'} and v['type'] in ('date', 'datetime', 'time'):
        return {'date': date, 'datetime': datetime, 'time': time}[v['type']].fromisoformat(v['value'])
    return v


def from_excel_value(v):
    if isinstance(v, (datetime, date, time)):
        return {'type': type(v).__name__, 'value': v.isoformat()}
    return v


def excel_cells(ws):
    out = {}
    for row in ws.iter_rows():
        for c in row:
            if c.value is None and not c.hyperlink and not c.comment:
                continue
            item = {'value': from_excel_value(c.value)}
            if c.hyperlink:
                item['hyperlink'] = {'target': c.hyperlink.target, 'location': c.hyperlink.location}
            if c.comment:
                item['note'] = {'text': c.comment.text, 'author': c.comment.author}
            out[(c.row, c.column)] = item
    return out


def json_cells(sheet):
    out = {}
    for r in sheet['rows']:
        for c in r['cells']:
            item = {'value': c.get('value')}
            if c.get('hyperlink'):
                item['hyperlink'] = c['hyperlink']
            if c.get('note'):
                item['note'] = c['note']
            out[(r['row'], c['column'])] = item
    return out


def diff_counts(ctx, master):
    wb = load_workbook(ctx.book)
    try:
        names = [s['name'] for s in master['sheets']]
        extra = [n for n in wb.sheetnames if n not in names]
        changed = 0
        for s in master['sheets']:
            if s['name'] not in wb.sheetnames:
                changed += len(json_cells(s))
                continue
            a, b = excel_cells(wb[s['name']]), json_cells(s)
            changed += sum(1 for k in set(a) | set(b) if a.get(k) != b.get(k))
        return changed, extra
    finally:
        wb.close()


def _show(v):
    if v is None or v == '':
        return '空'
    if isinstance(v, dict):
        v = v.get('value')
    v = str(v).replace('\n', ' ')
    return v if len(v) <= 60 else v[:57] + '…'


def _part(v):
    if not v:
        return None
    if isinstance(v, dict):
        return '｜'.join(f'{k}={v[k]}' for k in sorted(v) if v[k] not in (None, ''))
    return v


ROW_KEY_COLUMNS = {'校招主链接': ('公司',), '校招计划': ('公司', '计划'), '投递名额': ('公司', '范围')}


def _by_name(cells):
    """{(行,列): 单元格} → 表头名 与 {行号: {列名: 单元格}}；没有表头名的列记为“第N列”。"""
    head = {c: (v.get('value') if isinstance(v, dict) else v) for (r, c), v in cells.items() if r == 1}
    rows = {}
    for (r, c), v in cells.items():
        if r > 1:
            rows.setdefault(r, {})[head.get(c) or f'第{c}列'] = v
    return [h for _, h in sorted(head.items()) if h], rows


def _identity(ctx, sheet_name, head, row):
    """按内容认行：地区表用 公司+岗位链接/岗位名（公司行只按公司）；其他表用公司+计划/范围；兜底用第一列。"""
    val = lambda h: (row.get(h) or {}).get('value')
    if sheet_name in ctx.sheets and COMPANY_COL in head:
        link, name = val(LINK_COL), val(NAME_COL)
        k = {'sheet': sheet_name, 'company': val(COMPANY_COL)}
        if link:
            k['link'] = link
        if name:
            k['title'] = name
        if not link and not name:
            k['placeholder'] = True
        return key_id(k), describe(k)
    cols = ROW_KEY_COLUMNS.get(sheet_name) or tuple(head[:1])
    parts = [str(val(h) or '') for h in cols]
    return json.dumps([sheet_name, parts], ensure_ascii=False), f"{sheet_name}｜{'｜'.join(p for p in parts if p) or '空行'}"


def _keyed(ctx, sheet_name, head, rows):
    out, seen = {}, {}
    for rn in sorted(rows):
        k, label = _identity(ctx, sheet_name, head, rows[rn])
        n = seen[k] = seen.get(k, 0) + 1
        out[(k, n)] = (rn, label if n == 1 else f'{label}（第{n}条同名）', rows[rn])
    return out


def _row_status(ctx, row, head):
    """新增/删除的岗位行、公司行同时显示投递状态；空状态明确标出（须 propose→用户确认）。"""
    if ctx.column not in head:
        return ''
    v = (row.get(ctx.column) or {}).get('value')
    return f'｜投递状态：{v}' if v not in (None, '') else '｜投递状态：空（未确认，须 propose→用户确认）'


def diff_details(ctx, master, limit=30):
    """JSON 写入 Excel 后的实际变化：按内容（公司+岗位/计划/范围）配对行、按列名配对列比较，
    只列新增行、删除行和同一行里真正改变的单元格；行或列只是位置挪动的，写入后自动对齐，只报数量。"""
    wb = load_workbook(ctx.book)
    lines, by_col = [], {}
    added = removed = changed_rows = cells = moved = 0
    try:
        for s in master['sheets']:
            ehead, erows = _by_name(excel_cells(wb[s['name']]) if s['name'] in wb.sheetnames else {})
            jhead, jrows = _by_name(json_cells(s))
            name = s['name']
            gone, new_cols = [h for h in ehead if h not in jhead], [h for h in jhead if h not in ehead]
            if gone or new_cols:
                lines.append(f"{name}｜表头：" + '；'.join(([f"删除列：{'、'.join(gone)}"] if gone else [])
                                                      + ([f"新增列：{'、'.join(new_cols)}"] if new_cols else [])))
            if name not in wb.sheetnames:
                lines.append(f'{name}：新增整张工作表（{len(jrows)} 行）')
                added += len(jrows)
                continue
            old, cur = _keyed(ctx, name, ehead, erows), _keyed(ctx, name, jhead, jrows)
            if name not in ctx.sheets:
                # 计划名/范围改了：同一公司只有一条消失、一条新出现时，当作同一行的修改
                lost = {}
                for k in [k for k in old if k not in cur]:
                    lost.setdefault(json.loads(k[0])[1][0], []).append(k)
                for k in [k for k in cur if k not in old]:
                    comp = json.loads(k[0])[1][0]
                    fresh = [x for x in cur if x not in old and json.loads(x[0])[1][0] == comp]
                    if len(lost.get(comp, [])) == 1 and len(fresh) == 1:
                        old[k] = old.pop(lost.pop(comp)[0])
            for k in [k for k in cur if k not in old]:
                added += 1
                lines.append(f'新增：{cur[k][1]}{_row_status(ctx, cur[k][2], jhead)}（第{cur[k][0]}行）')
            for k in [k for k in old if k not in cur]:
                removed += 1
                lines.append(f'删除：{old[k][1]}{_row_status(ctx, old[k][2], ehead)}（原第{old[k][0]}行）')
            for k in [k for k in cur if k in old]:
                (orn, _, orow), (rn, label, row) = old[k], cur[k]
                moved += orn != rn
                items = []
                for col in [h for h in jhead if h in ehead] + [h for h in sorted(set(row) | set(orow)) if h.startswith('第') and h.endswith('列')]:
                    a, b = orow.get(col) or {}, row.get(col) or {}
                    if a == b:
                        continue
                    cells += 1
                    by_col[col] = by_col.get(col, 0) + 1
                    if a.get('value') != b.get('value'):
                        items.append(f"{col}：当前「{_show(a.get('value'))}」→写入「{_show(b.get('value'))}」")
                    else:
                        for part, zh in (('hyperlink', '超链接'), ('note', '批注')):
                            if a.get(part) != b.get(part):
                                items.append(f"{col}{zh}：当前「{_show(_part(a.get(part)))}」→写入「{_show(_part(b.get(part)))}」")
                if items:
                    changed_rows += 1
                    lines.append(f'修改：{label}（第{rn}行）：' + '；'.join(items))
    finally:
        wb.close()
    head_line = (f'实际变化：新增 {added} 行、删除 {removed} 行、修改 {changed_rows} 行（{cells} 个单元格）'
                 + (f'；另有 {moved} 行只是位置挪动（写入后自动对齐，内容不变）' if moved else ''))
    cols = '修改的列：' + '、'.join(f'{k} {v}' for k, v in sorted(by_col.items(), key=lambda x: -x[1])) if by_col else ''
    shown = lines[:limit]
    more = [f'……其余 {len(lines) - limit} 条未展开，可运行 json_excel_sync.py diff --limit 0 查看全部'] if len(lines) > limit else []
    return {'rows': added + removed + changed_rows, 'added': added, 'removed': removed, 'changed': changed_rows,
            'moved': moved, 'cells': cells, 'by_column': by_col,
            'text': '\n'.join([head_line] + ([cols] if cols else []) + shown + more)}

def excel_statuses(ctx):
    wb = load_workbook(ctx.book, read_only=True)
    try:
        current = {}
        for name in ctx.sheets:
            if name not in wb.sheetnames:
                continue
            rows = list(wb[name].iter_rows(values_only=True))
            if not rows:
                continue
            idx = {v: i + 1 for i, v in enumerate(rows[0]) if v is not None}
            for n, vals in enumerate(rows[1:], start=2):
                cells = [{'column': i + 1, 'value': v} for i, v in enumerate(vals) if v is not None]
                if not has_content(cells):
                    continue
                rk = row_key(name, {'row': n, 'cells': cells}, idx)
                st = (vals[idx[ctx.column] - 1] if ctx.column in idx else None) or ''
                current[key_id(rk)] = st
                if rk.get('link'):  # 改名后按“同表同公司同链接”找回原行；同链接多行且状态不一致时不作数
                    lk = link_key(rk)
                    current[lk] = st if lk not in current or current[lk] == st else None
        return current
    finally:
        wb.close()


def link_key(rk):
    return 'link:' + key_id({'sheet': rk['sheet'], 'company': rk['company'], 'link': rk['link']})


def identity_map(ctx):
    """relink 登记、尚未写入 Excel 的身份变更：新业务键 → 旧业务键（Excel 里这一行还是旧键）。"""
    return {key_id(c['new_key']): key_id(c['old_key'])
            for c in ctx.ledger().get('identity_changes', []) if not c.get('applied_at')}


def status_changes(ctx, master):
    """Excel 当前状态 → JSON 目标状态 的逐行差异。"""
    current = excel_statuses(ctx)
    renamed = identity_map(ctx)
    changes = []
    for name, s, idx, row in iter_status_rows(ctx, master):
        rk = row_key(name, row, idx)
        kid = key_id(rk)
        to = cell_value(row, idx[ctx.column]) or ''
        frm = current.get(kid)
        if frm is None and rk.get('link'):
            frm = current.get(link_key(rk))
        if frm is None and kid in renamed:  # relink 换过链接：按用户确认的旧业务键找回 Excel 中的原行
            frm = current.get(renamed[kid])
        frm = frm or ''
        if frm != to:
            changes.append({'key': json.loads(kid), 'from': frm, 'to': to})
    return changes


def unconfirmed(ctx, changes):
    open_entries = [e for e in ctx.ledger()['entries'] if not e.get('applied_at') and not e.get('superseded_at')]
    missing = []
    for ch in changes:
        if not any(key_id(e['key']) == key_id(ch['key']) and e['to'] == ch['to'] for e in open_entries):
            missing.append(ch)
    return missing


def apply_validation(ws, column_index, labels, max_row):
    letter = get_column_letter(column_index)
    keep = []
    for dv in ws.data_validations.dataValidation:
        if any(r.min_col == r.max_col == column_index for r in dv.sqref.ranges):
            continue
        keep.append(dv)
    ws.data_validations.dataValidation = keep
    formula = '"' + ','.join(labels) + '"'
    if len(formula) > 255:
        raise SystemExit('标签下拉列表超过 Excel 255 字符上限')
    dv = DataValidation(type='list', formula1=formula, allow_blank=False, showErrorMessage=True,
                        errorTitle='非法投递状态', error='只能选择标签体系中定义的投递状态')
    dv.add(f'{letter}2:{letter}{max(max_row, 2000)}')
    ws.add_data_validation(dv)


def write(ctx):
    master = ctx.load_master()
    problems = validate_labels(ctx, master) + validate_fields(ctx, master) + validate_quota(master) + validate_link_exceptions(ctx, master)
    if problems:
        raise SystemExit('标签/字段门禁拒绝写入（共 %d 处）：\n' % len(problems) + '\n'.join(problems[:50]))
    changes = status_changes(ctx, master)
    missing = unconfirmed(ctx, changes)
    if missing:
        lines = [f"{describe(m['key'])}：当前状态是「{m['from'] or '空'}」，更新状态是「{m['to']}」" for m in missing[:50]]
        raise SystemExit('状态变更门禁拒绝写入：以下 %d 行变更没有用户确认记录：\n' % len(missing) + '\n'.join(lines))
    wb = load_workbook(ctx.book)
    extra = [n for n in wb.sheetnames if n not in [s['name'] for s in master['sheets']]]
    if extra:
        raise SystemExit('Excel 含主数据没有的工作表，拒绝静默删除：' + '、'.join(extra))
    # Excel 由 JSON 派生，不做快照；出错时以主数据重新 write 即可恢复。
    for s in master['sheets']:
        ws = wb[s['name']] if s['name'] in wb.sheetnames else wb.create_sheet(s['name'])
        target, existing = json_cells(s), excel_cells(ws)
        for (r, c) in existing:
            if (r, c) not in target:
                cell = ws.cell(r, c)
                cell.value = None
                cell.hyperlink = None
                cell.comment = None
        last_styled = ws.max_row
        for (r, c), item in sorted(target.items()):
            cell = ws.cell(r, c)
            if not cell.has_style:
                if r > last_styled and r > 2:
                    cell._style = copy.copy(ws.cell(r - 1, c)._style)
                elif c > 1:
                    cell._style = copy.copy(ws.cell(r, c - 1)._style)
            v = to_excel_value(item['value'])
            cell.value = v
            if isinstance(v, str) and not v.startswith('='):
                cell.data_type = 's'
            if isinstance(v, (date, datetime)) and cell.number_format == 'General':
                cell.number_format = 'yyyy-mm-dd'
            h = item.get('hyperlink')
            if h and h.get('target'):
                cell.hyperlink = h['target']
            elif h and h.get('location'):
                cell.hyperlink = '#' + h['location']
            else:
                cell.hyperlink = None
            n = item.get('note')
            cell.comment = Comment(n['text'], n.get('author') or '') if n else None
        for t in s.get('tables', []):
            if t['name'] in ws.tables and ws.tables[t['name']].ref != t['ref']:
                wt = ws.tables[t['name']]
                wt.ref = t['ref']
                wt.tableColumns = []
                if wt.autoFilter is not None:
                    wt.autoFilter.ref = t['ref']
        if s['name'] in ctx.sheets:
            apply_validation(ws, headers(s)[ctx.column], list(ctx.labels), ws.max_row)
        if ws.auto_filter.ref:
            ws.auto_filter.ref = ws.dimensions
    buf = io.BytesIO()
    wb.save(buf)
    wb.close()
    write_bytes(ctx.book, buf.getvalue())
    changed, extra = diff_counts(ctx, master)
    if changed or extra:
        raise SystemExit(f'读回不一致：{changed} 个单元格与 JSON 不同；请以主数据重新 write')
    rb = load_workbook(ctx.book)
    try:
        formula = '"' + ','.join(ctx.labels) + '"'
        for name in ctx.sheets:
            letter = get_column_letter(headers(sheet_of(master, name))[ctx.column])
            if not any(dv.formula1 == formula and str(dv.sqref).startswith(letter + '2:') for dv in rb[name].data_validations.dataValidation):
                raise SystemExit(f'{name} 投递状态下拉校验读回失败')
    finally:
        rb.close()
    sha, applied = file_sha(ctx.book), now()
    final = {key_id(ch['key']): ch['to'] for ch in changes}
    ledger = ctx.ledger()
    for e in ledger['entries']:
        if e.get('applied_at') or e.get('superseded_at'):
            continue
        kid = key_id(e['key'])
        if kid in final and final[kid] == e['to']:
            e['applied_at'], e['applied_workbook_sha256'] = applied, sha
        else:
            e['superseded_at'] = applied
    for c in ledger.get('identity_changes', []):
        if not c.get('applied_at'):
            c['applied_at'] = applied
    ledger['deferrals'] = []
    save_json(ctx.ledger_path, ledger)
    refresh_digest(master)
    master['workbook_sha256'] = sha
    master['excel_written'] = {'data_sha256': master['data_sha256'], 'workbook_sha256': sha,
                               'written_at': applied}
    save_master(ctx, master)
    return {'written': True, 'status_changes_applied': len(changes), 'workbook_sha256': sha}


# ---------- 状态变更流程 ----------

def propose(ctx, items, master=None):
    master = master or ctx.load_master()
    out = []
    for d in items:
        key = normalize_key(d)
        s, idx, row = find_row(ctx, master, key)
        cur = cell_value(row, idx[ctx.column]) or ''
        to = d['to']
        err = label_problem(ctx, to, row_kind(row, idx))
        if err:
            raise SystemExit(f'标签门禁拒绝：{describe(key)} {err}')
        if to == cur:
            continue
        out.append({'key': key, 'from': cur, 'to': to,
                    'text': f"{describe(key)}：当前状态是「{cur or '空'}」，更新状态是「{to}」"})
    return out


QUOTA_SHEET = '投递名额'
QUOTA_HEAD = ['公司', '范围', '适用计划/批次', '投递上限', '当前投递数量', '是否分BU可平行投递', '是否可以滚动投递', '规则原文', '来源/核验日期']
TRI = ('是', '否', '公司未披露')
# 规则原文：官方来源文字的完整摘抄，写入时与来源原文逐字比对；没有官方说法时只能写“公司未披露”。
QUOTE_SOURCE_KEY = '原文出处文件'
# 来源/核验日期：来源：<公司>-<岗位或页面> 核验日期：YYYY-MM-DD
SOURCE_FORMAT = '来源：<公司>-<岗位或页面> 核验日期：YYYY-MM-DD'


def _norm_quote(text):
    import unicodedata
    return re.sub(r'\s+', '', unicodedata.normalize('NFKC', str(text)))


def quota_source_problem(company, value):
    m = re.fullmatch(r'来源：' + re.escape(str(company)) + r'\s*-\s*(\S.*?) 核验日期：(\d{4}-\d{2}-\d{2})', str(value or ''))
    if not m:
        return f'来源/核验日期格式必须为“{SOURCE_FORMAT}”，且公司与本行一致：{value!r}'
    try:
        d = date.fromisoformat(m.group(2))
    except ValueError:
        return f'核验日期不是有效日期：{m.group(2)}'
    if d > date.today():
        return f'核验日期晚于今天：{m.group(2)}'
    return None


def quota_quote_problem(item):
    """规则原文必须是来源原文的连续摘抄（忽略空白与全/半角差异），不能总结、改写或拼接多段。"""
    text = item.get('规则原文')
    if text == '公司未披露':
        return None
    path = item.get(QUOTE_SOURCE_KEY)
    if not path:
        return f'缺少“{QUOTE_SOURCE_KEY}”：须提供保存的官方来源原文（UTF-8 文本），用于逐字核对规则原文'
    try:
        src = Path(path).read_text(encoding='utf-8')
    except (OSError, UnicodeDecodeError) as e:
        return f'来源原文文件不可读：{path}（{e}）'
    q = _norm_quote(text)
    if not q:
        return '规则原文不能为空'
    if q not in _norm_quote(src):
        return '规则原文不是来源原文的逐字连续摘抄（不能总结、改写或用“；”拼接多段；多条规则请按范围分行写）'
    return None


def quota_sheet(master, create=False):
    for sh in master['sheets']:
        if sh['name'] == QUOTA_SHEET:
            return sh
    if not create:
        return None
    sh = {'name': QUOTA_SHEET, 'rows': [{'row': 1, 'cells': [{'column': i + 1, 'value': h, 'type': 's'} for i, h in enumerate(QUOTA_HEAD)]}],
          'max_row': 1, 'max_column': len(QUOTA_HEAD), 'state': 'visible', 'merged_ranges': [], 'tables': []}
    master['sheets'].append(sh)
    return sh


def quota_records(master):
    sh = quota_sheet(master)
    if not sh:
        return []
    idx = headers(sh)
    return [(r, {h: cell_value(r, idx[h]) for h in QUOTA_HEAD if h in idx}) for r in sh['rows'] if r['row'] > 1]


def validate_quota(master):
    probs, seen = [], {}
    for r, v in quota_records(master):
        k = (v.get('公司'), v.get('范围'))
        where = f"{QUOTA_SHEET}!{r['row']} {v.get('公司')}"
        if not v.get('公司') or not v.get('范围'):
            probs.append(where + ' 公司/范围不能为空')
        if k in seen:
            probs.append(where + f' 与第{seen[k]}行重复：同一公司同一范围只能有一条规则')
        seen[k] = r['row']
        lim = v.get('投递上限')
        if not (isinstance(lim, int) and lim >= 0) and lim != '公司未披露':
            probs.append(where + f' 投递上限只能是非负整数或“公司未披露”：{lim!r}')
        for f in ('是否分BU可平行投递', '是否可以滚动投递'):
            if v.get(f) not in TRI:
                probs.append(where + f' {f}只能是 是/否/公司未披露：{v.get(f)!r}')
        for f in ('适用计划/批次', '规则原文', '来源/核验日期'):
            if v.get(f) in (None, ''):
                probs.append(where + f' {f}不能为空')
    return probs


def scope_rows(ctx, master, company, scope=None, plan=None):
    out = []
    for name in ctx.sheets:
        s = sheet_of(master, name)
        idx = headers(s)
        for r in s['rows']:
            if r['row'] == 1 or row_kind(r, idx) != 'role' or cell_value(r, idx[COMPANY_COL]) != company:
                continue
            if scope not in (None, '公司总量'):
                bu = cell_value(r, idx['所属BU']) if '所属BU' in idx else None
                pl = cell_value(r, idx['校招计划']) if '校招计划' in idx else None
                if not ((bu and bu == scope) or (pl and plan and pl == plan)):
                    continue
            out.append((s, idx, r))
    return out


def refresh_quota_counts(ctx, master):
    """当前投递数量是派生列：每次保存按投递状态重新计算，不允许手填。"""
    sh = quota_sheet(master)
    rule = ctx.cfg.get('quotaRule') or {}
    if not sh or not rule:
        return
    counts = set(rule.get('countsStatuses', []))
    col = headers(sh)['当前投递数量']
    for r, v in quota_records(master):
        rows = scope_rows(ctx, master, v['公司'], v['范围'], v.get('适用计划/批次'))
        set_cell(r, col, sum(1 for s_, idx, x in rows if cell_value(x, idx[ctx.column]) in counts))


def set_quota_rules(ctx, items):
    """按公司覆盖写名额表：输入里出现的公司，其旧规则整体删除，换成这一版（只保留最新）。"""
    master = ctx.load_master()
    sh = quota_sheet(master, create=True)
    companies = {d['公司'] for d in items}
    keep = [r for r in sh['rows'] if r['row'] == 1 or cell_value(r, headers(sh)['公司']) not in companies]
    bad = []
    for i, d in enumerate(items, start=1):
        unknown = set(d) - set(QUOTA_HEAD) - {QUOTE_SOURCE_KEY}
        if unknown or '当前投递数量' in d:
            raise SystemExit('名额表字段非法（当前投递数量由脚本计算）：' + '、'.join(sorted(unknown | ({'当前投递数量'} & set(d)))))
        where = f"第{i}条 {d.get('公司')}｜{d.get('范围')}："
        for prob in (quota_source_problem(d.get('公司'), d.get('来源/核验日期')), quota_quote_problem(d)):
            if prob:
                bad.append(where + prob)
    if bad:
        raise SystemExit('名额表门禁拒绝写入（共 %d 处）：\n' % len(bad) + '\n'.join(bad[:50]))
    for d in items:
        keep.append({'row': 0, 'cells': [{'column': QUOTA_HEAD.index(h) + 1, 'value': d[h], 'type': 'n' if isinstance(d[h], int) else 's'}
                                          for h in QUOTA_HEAD if d.get(h) not in (None, '')]})
    head, body = keep[0], keep[1:]
    body.sort(key=lambda r: (str(cell_value(r, 1)), str(cell_value(r, 2))))
    for i, r in enumerate(body, start=2):
        r['row'] = i
    sh['rows'] = [head] + body
    sh['max_row'] = len(sh['rows'])
    probs = validate_quota(master)
    if probs:
        raise SystemExit('名额表门禁拒绝写入（共 %d 处）：\n' % len(probs) + '\n'.join(probs[:50]))
    save_master(ctx, master)
    return {'companies': len(companies), 'rows': len(body)}


def set_sheet_rows(ctx, sheet_name, rows):
    """整页覆盖说明类工作表（如“分类说明”）：只保留最新一版。地区岗位表与“投递名额”不能用此命令。"""
    if sheet_name in ctx.sheets or sheet_name == QUOTA_SHEET:
        raise SystemExit('岗位表与名额表只能用对应的固定命令写入')
    if not rows or not all(isinstance(r, list) and r for r in rows) or any(v in (None, '') for v in rows[0]):
        raise SystemExit('输入须为二维列表，首行为非空表头')
    master = ctx.load_master()
    s = sheet_of(master, sheet_name)
    s['rows'] = [{'row': i, 'cells': [{'column': j, 'value': v, 'type': 'n' if isinstance(v, (int, float)) else 's'}
                                        for j, v in enumerate(r, start=1) if v not in (None, '')]}
                 for i, r in enumerate(rows, start=1)]
    s['max_row'] = len(rows)
    s['max_column'] = max(len(r) for r in rows)
    for t in s.get('tables', []):
        t['ref'] = 'A1:%s%d' % (get_column_letter(s['max_column']), len(rows))
    save_master(ctx, master)
    return {'sheet': sheet_name, 'rows': len(rows)}


def drop_column(ctx, sheet_name, column_name):
    """删除整列（后续列左移），用于去掉与名额表重复的规则列；表格范围同步收窄。"""
    master = ctx.load_master()
    s = sheet_of(master, sheet_name)
    idx = headers(s)
    if column_name not in idx:
        return {'dropped': False}
    if column_name in (COMPANY_COL, NAME_COL, LINK_COL, ctx.column):
        raise SystemExit('不能删除业务键或投递状态列')
    col = idx[column_name]
    for r in s['rows']:
        r['cells'] = [c for c in r['cells'] if c['column'] != col]
        for c in r['cells']:
            if c['column'] > col:
                c['column'] -= 1
    s['max_column'] = max(1, s.get('max_column', col) - 1)
    for t in s.get('tables', []):
        a, b = t['ref'].split(':')
        end = re.match(r'([A-Z]+)(\d+)', b)
        t['ref'] = f"{a}:{get_column_letter(max(1, column_index_from_string(end.group(1)) - 1))}{end.group(2)}"
    save_master(ctx, master)
    return {'dropped': True, 'sheet': sheet_name, 'column': column_name}


def load_quotas(ctx, master):
    rule = ctx.cfg.get('quotaRule')
    if not rule:
        return None, []
    qs = [{'company': v['公司'], 'scope': v['范围'], 'plan': v.get('适用计划/批次'), 'limit': v['投递上限']}
          for r, v in quota_records(master) if v.get('范围') == '公司总量' and isinstance(v.get('投递上限'), int)]
    return rule, qs


def quota_closure(ctx, master):
    """名额被用户选择填满后，同一名额范围内剩余的待筛选岗位 → 不符合用户要求。返回派生变更。"""
    rule, quotas = load_quotas(ctx, master)
    derived, over = [], []
    if not rule:
        return derived, over
    counts, sources, target = set(rule['countsStatuses']), set(rule['autoCloseFrom']), rule['autoCloseTo']
    for q in quotas:
        scope = scope_rows(ctx, master, q['company'])
        used = [r for s, idx, r in scope if cell_value(r, idx[ctx.column]) in counts]
        if len(used) > q['limit']:
            over.append(f"{q['company']}：已占用 {len(used)} 个，超过名额 {q['limit']}")
        if len(used) >= q['limit']:
            for s, idx, r in scope:
                cur = cell_value(r, idx[ctx.column])
                if cur in sources:
                    key = row_key(s['name'], r, idx)
                    derived.append({'key': key, 'from': cur, 'to': target, 'quota': q,
                                    'text': f"{describe(key)}：当前状态是「{cur}」，更新状态是「{target}」"
                                            f"（名额收口：本轮名额 {q['limit']} 已由用户选择占满）"})
    return derived, over


def confirm(ctx, items, quote, ref):
    if not quote or not quote.strip():
        raise SystemExit('必须提供用户确认原话')
    master = ctx.load_master()
    plan = propose(ctx, items, master)
    ledger = ctx.ledger()
    excel_now = excel_statuses(ctx)
    at = now()
    for p in plan:
        s, idx, row = find_row(ctx, master, p['key'])
        set_cell(row, idx[ctx.column], p['to'])
        ledger['entries'].append({'id': uuid.uuid4().hex, 'key': p['key'], 'from': p['from'],
                                  'from_excel': excel_now.get(key_id(p['key']), ''), 'to': p['to'],
                                  'shown_to_user': p['text'], 'user_quote': quote.strip(),
                                  'confirmation_ref': ref, 'confirmed_at': at, 'applied_at': None})
    derived, over = quota_closure(ctx, master)
    if over:
        raise SystemExit('超过本轮名额，拒绝确认（未写入任何文件）：' + '；'.join(over))
    for p in derived:
        s, idx, row = find_row(ctx, master, p['key'])
        set_cell(row, idx[ctx.column], p['to'])
        ledger['entries'].append({'id': uuid.uuid4().hex, 'key': p['key'], 'from': p['from'],
                                  'from_excel': excel_now.get(key_id(p['key']), ''), 'to': p['to'],
                                  'shown_to_user': p['text'], 'user_quote': quote.strip(),
                                  'confirmation_ref': ref, 'confirmed_at': at, 'applied_at': None,
                                  'derived_by': 'quota_closure', 'quota': p['quota']})
    refresh_digest(master)
    save_master(ctx, master)
    save_json(ctx.ledger_path, ledger)
    return {'confirmed': len(plan), 'quota_closure_changes': [p['text'] for p in derived], 'excel_written': False,
            'next': '对话结束前询问用户：现在写入 Excel，还是下一轮再写入'}


def upsert(ctx, items):
    master = ctx.load_master()
    added = updated = 0
    duplicates_removed = 0
    touched = []
    for d in items:
        key = normalize_key(d['key'])
        fields = d['values']
        if ctx.column in fields:
            raise SystemExit('upsert 不能修改投递状态；状态变更必须走 propose→用户确认→confirm')
        s = sheet_of(master, key['sheet'])
        idx = headers(s)
        unknown = set(fields) - set(idx)
        if unknown:
            raise SystemExit('未知表头：' + '、'.join(sorted(unknown)))
        hits = [r for r in s['rows'] if r['row'] > 1 and all(
            row_key(key['sheet'], r, idx).get(k) == v for k, v in key.items())]
        if len(hits) > 1 and d.get('drop_empty_duplicates'):
            if not key.get('link') or not key.get('title'):
                raise SystemExit('清理空状态重复行必须同时提供岗位链接和岗位名称')
            kept = [r for r in hits if cell_value(r, idx[ctx.column]) not in (None, '')]
            if len(kept) != 1 or any(row_key(key['sheet'], r, idx) != key for r in hits):
                raise SystemExit('只能清理同一完整业务键下、唯一有状态岗位旁的空状态重复行')
            gone = [r for r in hits if r is not kept[0]]
            removed = sorted(r['row'] for r in gone)
            gone_ids = {id(r) for r in gone}
            s['rows'] = [r for r in s['rows'] if id(r) not in gone_ids]
            for r in s['rows']:
                r['row'] -= sum(n < r['row'] for n in removed)
            s['max_row'] = max(r['row'] for r in s['rows'])
            duplicates_removed += len(gone)
            hits = kept
        if len(hits) > 1:
            raise SystemExit('业务键匹配多行：' + key_id(key))
        if hits:
            row = hits[0]
            updated += 1
        else:
            # 公司行固定保留：新岗位不占用公司行，排在该公司最后；新公司行排在该公司最前
            company_rows = [r for r in s['rows'] if r['row'] > 1 and cell_value(r, idx[COMPANY_COL]) == key['company']]
            if key.get('placeholder') and company_rows:
                at = min(r['row'] for r in company_rows)
            elif company_rows:
                at = max(r['row'] for r in company_rows) + 1
            else:
                at = max(r['row'] for r in s['rows']) + 1
            row = insert_row(s, idx, at, company_rows[0] if company_rows else None)
            added += 1
            for col, v in ((COMPANY_COL, key['company']), (LINK_COL, key.get('link')), (NAME_COL, key.get('title'))):
                if v and col in idx:
                    set_cell(row, idx[col], v)
        for col, v in fields.items():
            if col == NOTE_COL and NOTE_COL in idx:
                cur_note = str(cell_value(row, idx[NOTE_COL]) or '')
                if (str(v or '').startswith(LINK_EXCEPTION_PREFIX) or cur_note.startswith(LINK_EXCEPTION_PREFIX)) and str(v or '') != cur_note:
                    raise SystemExit(f'upsert 不能写入或修改链接例外备注：{describe(key)}。“{LINK_EXCEPTION_PREFIX}”开头的备注只能用 '
                                     'link-exception（先 --preview 给用户看，确认后带 --quote 执行）')
            if (col == LINK_COL and hits and v not in (None, '')
                    and str(v) != str(cell_value(row, idx[LINK_COL]) or '')):
                raise SystemExit(f'upsert 不能更换已有岗位的链接：{describe(key)}。岗位链接是行身份的一部分，'
                                 '请用 relink（先 --preview 给用户看“岗位链接 X → Y”，确认后带 --quote 执行）')
            if isinstance(v, str) and len(v) > 32767:
                raise SystemExit('禁止截断单元格')
            if col in ctx.field_rules and v not in (None, ''):
                kind = 'role' if (key.get('link') or key.get('title')) else 'company'
                err = check_field(ctx, col, v, kind)
                if err:
                    raise SystemExit(f'字段门禁：{describe(key)} {err}')
            if col == LINK_COL and v in (None, ''):
                row['cells'] = [c for c in row['cells'] if c['column'] != idx[col]]  # 连同超链接一起清除
                continue
            set_cell(row, idx[col], v)
        link = cell_value(row, idx[LINK_COL]) if LINK_COL in idx else None
        if link:
            c = next(c for c in row['cells'] if c['column'] == idx[LINK_COL])
            c['hyperlink'] = {'target': link, 'location': None}
        touched.append((s, idx, row, key))
    errors = []
    for s, idx, row, key in touched:
        kind = row_kind(row, idx)
        for field in ctx.field_rules:
            if field in idx and (field in (NAME_COL, LINK_COL) or cell_value(row, idx[field]) not in (None, '')):
                err = check_field(ctx, field, cell_value(row, idx[field]), kind)
                if err:
                    errors.append(f'{describe(key)} {err}')
    if errors:
        raise SystemExit('字段门禁拒绝写入主数据（共 %d 处）：\n' % len(errors) + '\n'.join(errors[:50]))
    created = []
    for s, idx, row, key in touched:
        if ensure_company_row(s, idx, key['company']):
            created.append(f"{s['name']}｜{key['company']}")
    refresh_digest(master)
    save_master(ctx, master)
    return {'added': added, 'updated': updated, 'duplicates_removed': duplicates_removed, 'company_rows_created': created,
            'note': '新增行（含自动补的公司行）投递状态为空：必须向用户展示“当前状态是空，更新状态是Y”并确认后，才能写入 Excel'}


def insert_row(s, idx, at, template=None):
    """在第 at 行插入空行（其后各行下移），并从 template 复制公司信息列。"""
    for r in s['rows']:
        if r['row'] >= at:
            r['row'] += 1
    row = {'row': at, 'cells': []}
    if template is not None:
        for col in COMPANY_INFO_COLS:
            if col in idx:
                set_cell(row, idx[col], cell_value(template, idx[col]))
    s['rows'].append(row)
    s['rows'].sort(key=lambda r: r['row'])
    s['max_row'] = max(r['row'] for r in s['rows'])
    return row


def ensure_company_row(s, idx, company):
    """方案B：公司在该表有行但缺公司行时，在该公司最前面补一行公司行（状态为空，须经用户确认）。返回是否新建。"""
    rows = [r for r in s['rows'] if r['row'] > 1 and cell_value(r, idx[COMPANY_COL]) == company]
    if not rows or any(row_kind(r, idx) == 'company' for r in rows):
        return False
    first = min(rows, key=lambda r: r['row'])
    insert_row(s, idx, first['row'], first)
    return True


HANDED_OFF = ('已交接', '进行中')


def link_exception(ctx, items, quote, preview=False):
    """登记或撤销链接例外（官网没有独立的岗位详情页）。备注写成固定格式「用户例外保留：官网无独立详情页。<定位说明>」，
    forward 据此不检查该岗位链接是否为具体岗位页（仍须 https 官网、带官方岗位ID、不与同公司其它岗位共用）。
    preview 只返回给用户看的逐行变更；用户确认后带原话执行，确认记入台账（每行只留最新一条）。
    items：[{sheet,company,link,title,定位说明}]；撤销时给 "撤销": true。"""
    if not preview and (not quote or not quote.strip()):
        raise SystemExit('必须提供用户确认原话')
    master = ctx.load_master()
    plan, lines = [], []
    for d in items:
        key = normalize_key({k: d.get(k) for k in ('sheet', 'company', 'link', 'title')})
        if key.get('placeholder'):
            raise SystemExit(f'{describe(key)} 是公司行，不能登记链接例外')
        s, idx, row = find_row(ctx, master, key)
        if NOTE_COL not in idx:
            raise SystemExit(f"{s['name']} 没有“{NOTE_COL}”列")
        cur = str(cell_value(row, idx[NOTE_COL]) or '')
        if d.get('撤销'):
            if not cur.startswith(LINK_EXCEPTION_PREFIX):
                raise SystemExit(f'{describe(key)} 当前备注不是链接例外，无需撤销')
            new, effect = None, '生成投递批次时恢复检查该岗位链接是否为具体岗位页'
        else:
            how = str(d.get('定位说明') or '').strip()
            if len(how) < 4:
                raise SystemExit(f'{describe(key)} 须写明定位说明：官网上如何找到这个岗位（如“列表页按 positionId 定位，点岗位名展开 JD”）')
            if cur and not cur.startswith(LINK_EXCEPTION_PREFIX):
                raise SystemExit(f'{describe(key)} 已有其它用途的备注「{cur}」，备注只能放一条，请先与用户确认保留哪一条')
            new, effect = LINK_EXCEPTION_PREFIX + how, '生成投递批次时不检查该岗位链接是否为具体岗位页'
            err = check_field(ctx, NOTE_COL, new, 'role')
            if err:
                raise SystemExit(f'字段门禁：{describe(key)} {err}')
        line = f"{describe(key)}：备注「{cur or '空'}」→「{new or '空'}」（{effect}）"
        plan.append((s, idx, row, key, new, line))
        lines.append(line)
    if preview:
        return {'show_user_exactly': lines,
                'rule': '必须把 show_user_exactly 原样展示给用户并取得明确确认后，才能用同一 input 加 --quote 执行'}
    ledger = ctx.ledger()
    records = ledger.setdefault('link_exceptions', [])
    at = now()
    for s, idx, row, key, new, line in plan:
        set_cell(row, idx[NOTE_COL], new)
        records[:] = [e for e in records if key_id(e['key']) != key_id(key)]
        if new:
            records.append({'id': uuid.uuid4().hex, 'key': key, 'note': new, 'shown_to_user': line,
                            'user_quote': quote.strip(), 'confirmed_at': at})
    refresh_digest(master)
    save_master(ctx, master)
    save_json(ctx.ledger_path, ledger)
    return {'link_exceptions': len(plan), 'excel_written': False,
            'next': '对话结束前询问用户：现在写入 Excel，还是下一轮再写入'}


def relink(ctx, items, quote, preview=False):
    """更换已有岗位行的岗位链接。岗位链接是行身份的一部分，所以换链接不走 upsert：
    preview 只返回给用户看的逐行“岗位链接 X → Y”，不改任何文件；用户确认后带原话执行——
    改链接与超链接；把该行的确认记录迁到新业务键；登记身份变更（写 Excel 时据此认出同一行，状态不变、无需重新确认）；
    把 公司.json 里的行身份迁到新链接（志愿顺序、城市选项等随行保留）。
    该岗位在已交接/进行中的投递批次里时拒绝，须先完成投递与反向评审。"""
    if not preview and (not quote or not quote.strip()):
        raise SystemExit('必须提供用户确认原话')
    master = ctx.load_master()
    store = ctx.cfg.get('companyStore') or {}
    batch_file = ctx.root / store.get('batchFile', '岗位交接/投递批次.json')
    batch = json.loads(batch_file.read_text(encoding='utf-8')) if batch_file.exists() else None
    plan, lines, taken = [], [], set()
    for d in items:
        old_key = normalize_key({k: d.get(k) for k in ('sheet', 'company', 'link', 'title')})
        if old_key.get('placeholder'):
            raise SystemExit(f'{describe(old_key)} 是公司行，没有岗位链接可换')
        s, idx, row = find_row(ctx, master, old_key)
        if LINK_COL not in idx:
            raise SystemExit(f"{s['name']} 没有“{LINK_COL}”列")
        old = str(cell_value(row, idx[LINK_COL]) or '')
        new = str(d.get('new_link') or '').strip()
        if not re.fullmatch(r'https://\S+', new):
            raise SystemExit(f'{describe(old_key)} 新链接必须是完整的 https 链接：{new!r}')
        if new.rstrip('/') == old.rstrip('/'):
            raise SystemExit(f'{describe(old_key)} 新旧链接相同')
        norm = (s['name'], old_key['company'], new.rstrip('/'))
        clash = [r['row'] for r in s['rows'] if r['row'] > 1 and r is not row
                 and cell_value(r, idx[COMPANY_COL]) == old_key['company']
                 and str(cell_value(r, idx[LINK_COL]) or '').rstrip('/') == new.rstrip('/')]
        if clash or norm in taken:
            raise SystemExit(f'{describe(old_key)} 新链接已被同公司其它行使用' + (f'（第{clash[0]}行）' if clash else '（本次输入重复）'))
        taken.add(norm)
        status = cell_value(row, idx[ctx.column]) or ''
        in_batch = bool(batch) and old and any(
            c.get('公司') == old_key['company'] and str(r.get('岗位链接') or '').rstrip('/') == old.rstrip('/')
            for c in batch.get('companies', []) for r in c.get('岗位', []))
        if in_batch and batch.get('state') in HANDED_OFF:
            raise SystemExit(f"{describe(old_key)} 在已交接的投递批次 {batch.get('batchId')} 中（{batch.get('state')}），"
                             '须先完成投递与反向评审再换链接')
        new_key = {**old_key, 'link': new}
        line = (f"{describe(old_key)}：岗位链接「{old or '空'}」→「{new}」"
                f"（投递状态「{status or '空'}」保持不变，志愿顺序等投递数据随行迁移）")
        if in_batch:
            line += f"；该岗位在待评审批次 {batch.get('batchId')} 中，换链接后须重新生成批次（forward）"
        plan.append((s, idx, row, old_key, new_key, new, line))
        lines.append(line)
    if preview:
        return {'show_user_exactly': lines,
                'rule': '必须把 show_user_exactly 原样展示给用户并取得明确确认后，才能用同一 input 加 --quote 执行'}
    ledger = ctx.ledger()
    changes = ledger.setdefault('identity_changes', [])
    at = now()
    for s, idx, row, old_key, new_key, new, line in plan:
        set_cell(row, idx[LINK_COL], new)
        cell = next(c for c in row['cells'] if c['column'] == idx[LINK_COL])
        cell['hyperlink'] = {'target': new, 'location': None}
        ok, nk = key_id(old_key), key_id(new_key)
        for e in ledger['entries'] + ledger.get('link_exceptions', []):
            if key_id(e['key']) == ok:
                e['key'] = new_key
        prior = next((c for c in changes if not c.get('applied_at') and key_id(c['new_key']) == ok), None)
        if prior and key_id(prior['old_key']) == nk:  # 改回 Excel 里的原链接：不再需要身份变更
            changes.remove(prior)
        elif prior:  # 写 Excel 前连续换了两次：仍从 Excel 里的原键直接指向最新键
            prior.update({'new_key': new_key, 'shown_to_user': line, 'user_quote': quote.strip(), 'confirmed_at': at})
        else:
            changes.append({'id': uuid.uuid4().hex, 'old_key': old_key, 'new_key': new_key, 'shown_to_user': line,
                            'user_quote': quote.strip(), 'confirmed_at': at, 'applied_at': None})
    moved = []
    base = ctx.root / store.get('dir', '岗位交接/公司')
    if base.exists():
        remap = {key_id(o): n for _, _, _, o, n, _, _ in plan}
        for cf in sorted(base.glob('*/*/公司.json')):
            doc = json.loads(cf.read_text(encoding='utf-8'))
            hit = False
            for r in doc.get('roles', []):
                nk = remap.get(key_id(r.get('主表行身份') or {}))
                if nk:
                    r['主表行身份'] = nk
                    hit = True
                    moved.append(describe(nk))
            if hit:
                doc['updatedAt'] = at
                save_json(cf, doc)
    probs = validate_labels(ctx, master)
    if probs:
        raise SystemExit('标签门禁拒绝：' + '；'.join(probs[:10]))
    refresh_digest(master)
    save_master(ctx, master)
    save_json(ctx.ledger_path, ledger)
    return {'relinked': len(plan), 'company_files_updated': moved, 'excel_written': False,
            'next': '对话结束前询问用户：现在写入 Excel，还是下一轮再写入'}


KEEP_STATUSES = ('岗位已投递', '岗位推进中', '岗位offer', '岗位终止')


def reset_rows(ctx, items, quote, preview=False, master=None, drop_companies=(), mark_research=False):
    """删除用户指定的岗位行（如编造/无效的行）。公司行固定保留：公司已无岗位行时公司行改回“待搜索”并删掉该公司的名额规则；
    --mark-research 时即使仍有岗位也把公司行改回“待搜索”（保留名额规则）；--drop-company 删光后连同公司行整家删除。
    有真实投递记录的行（已投递/推进中/offer/终止）不能删。preview=True 时只返回要展示给用户的逐行变更。"""
    if not preview and (not quote or not quote.strip()):
        raise SystemExit('必须提供用户确认原话')
    master = master if master is not None else ctx.load_master()
    lines, dropped, placeholders = [], [], []
    by_sheet = {}
    for d in items:
        key = normalize_key(d)
        s, idx, row = find_row(ctx, master, key)
        st = cell_value(row, idx[ctx.column]) or ''
        if st in KEEP_STATUSES:
            raise SystemExit(f'{describe(key)} 状态为「{st}」，有真实投递记录，不能删除')
        by_sheet.setdefault(s['name'], []).append((key, row))
        lines.append(f'{describe(key)}：当前状态是「{st or "空"}」，更新为「删除该行」')
    resets = []  # 已有公司行改回“待搜索”：(sheet, row, key, from, keepRules)
    for name, pairs in by_sheet.items():
        s = sheet_of(master, name)
        idx = headers(s)
        gone = {id(r) for _, r in pairs}
        for company in dict.fromkeys(k['company'] for k, _ in pairs):
            rest = [r for r in s['rows'] if r['row'] > 1 and id(r) not in gone and cell_value(r, idx[COMPANY_COL]) == company]
            crow = next((r for r in rest if row_kind(r, idx) == 'company'), None)
            roles_left = [r for r in rest if row_kind(r, idx) == 'role']
            ckey = {'sheet': name, 'company': company, 'placeholder': True}
            if not roles_left and company in drop_companies:
                if crow is not None:
                    gone.add(id(crow))
                placeholders.append((s, None, {'sheet': name, 'company': company, 'dropped': True}))
                lines.append(f'{name}｜{company}：整家删除（含公司行），并删除该公司名额规则')
            elif mark_research or not roles_left:
                note = '（该公司需重新搜索；保留其余岗位与名额规则）' if roles_left else '（公司退回待搜索，并删除该公司名额规则）'
                if crow is None:  # 旧数据缺公司行：补一行
                    first = next(r for k, r in pairs if k['company'] == company)
                    ph = {'row': first['row'], 'cells': [dict(c) for c in first['cells'] if c['column'] in
                                                        [idx[h] for h in COMPANY_INFO_COLS if h in idx]]}
                    set_cell(ph, idx[ctx.column], '待搜索')
                    placeholders.append((s, ph, {**ckey, 'keepRules': bool(roles_left)}))
                    lines.append(f'{name}｜{company}｜公司行：当前状态是「空」，更新状态是「待搜索」{note}')
                else:
                    cur = cell_value(crow, idx[ctx.column]) or ''
                    resets.append((s, crow, ckey, cur, bool(roles_left)))
                    lines.append(f'{name}｜{company}｜公司行：当前状态是「{cur or "空"}」，更新状态是「待搜索」{note}')
        dropped.append((s, gone))
    if preview:
        return {'show_user_exactly': lines}
    deleted_keys = {key_id(k) for ps in by_sheet.values() for k, _ in ps}
    deleted_keys |= {key_id({'sheet': k['sheet'], 'company': k['company'], 'placeholder': True})
                     for _, ph, k in placeholders if ph is None}
    for s, crow, ckey, cur, keep in resets:
        set_cell(crow, headers(s)[ctx.column], '待搜索')
    for s, gone in dropped:
        phs = [ph for ss, ph, _ in placeholders if ss is s and ph is not None]
        removed = sorted({r['row'] for r in s['rows'] if id(r) in gone} - {ph['row'] for ph in phs})
        keep = [r for r in s['rows'] if id(r) not in gone] + phs
        for r in keep:  # 后续行上移，不改变其它行的相对位置
            r['row'] -= sum(1 for x in removed if x < r['row'])
        keep.sort(key=lambda r: r['row'])
        s['rows'] = keep
        s['max_row'] = max(r['row'] for r in keep)
    qs = quota_sheet(master)
    if qs and (placeholders or resets):
        drop = {k['company'] for _, _, k in placeholders if not k.get('keepRules')} | {k['company'] for _, _, k, _, keep in resets if not keep}
        qi = headers(qs)
        qs['rows'] = [r for r in qs['rows'] if r['row'] == 1 or cell_value(r, qi['公司']) not in drop]
        for i, r in enumerate(qs['rows'], start=1):
            r['row'] = i
        qs['max_row'] = len(qs['rows'])
    ledger = ctx.ledger()
    ledger['entries'] = [e for e in ledger['entries'] if key_id(e['key']) not in deleted_keys]  # 行已删除，旧确认不再保留
    if 'link_exceptions' in ledger:
        ledger['link_exceptions'] = [e for e in ledger['link_exceptions'] if key_id(e['key']) not in deleted_keys]
    at = now()
    for _, ph, key in placeholders:
        if ph is None:
            continue
        key = {k: v for k, v in key.items() if k != 'keepRules'}
        ledger['entries'] = [e for e in ledger['entries'] if key_id(e['key']) != key_id(key)]
        ledger['entries'].append({'id': str(uuid.uuid4()), 'key': key, 'from': '', 'from_excel': '', 'to': '待搜索',
                                  'shown_to_user': '公司退回待搜索', 'user_quote': quote.strip(), 'confirmation_ref': 'reset-rows', 'confirmed_at': at})
    for s, crow, key, cur, keep in resets:
        ledger['entries'] = [e for e in ledger['entries'] if key_id(e['key']) != key_id(key)]
        ledger['entries'].append({'id': str(uuid.uuid4()), 'key': key, 'from': cur, 'from_excel': '', 'to': '待搜索',
                                  'shown_to_user': f"{describe(key)}：当前状态是「{cur or '空'}」，更新状态是「待搜索」",
                                  'user_quote': quote.strip(), 'confirmation_ref': 'reset-rows', 'confirmed_at': at})
    probs = validate_labels(ctx, master)
    if probs:
        raise SystemExit('标签门禁拒绝：' + '；'.join(probs[:10]))
    save_master(ctx, master)
    save_json(ctx.ledger_path, ledger)
    return {'deleted_rows': sum(len(g) for _, g in dropped),
            'companies_reset': [k['company'] for _, ph, k in placeholders if ph is not None] + [k['company'] for _, _, k, _, _ in resets],
            'companies_dropped': [k['company'] for _, ph, k in placeholders if ph is None]}


def clear_orphans(ctx, rows, quote):
    """清空用户确认的“无公司名”行（只允许公司列为空的行），记入台账。"""
    if not quote or not quote.strip():
        raise SystemExit('必须提供用户确认原话')
    master = ctx.load_master()
    ledger = ctx.ledger()
    cleared = []
    for d in rows:
        s = sheet_of(master, d['sheet'])
        idx = headers(s)
        row = next((r for r in s['rows'] if r['row'] == d['row']), None)
        if row is None:
            raise SystemExit(f"{d['sheet']}!{d['row']} 不存在")
        if cell_value(row, idx[COMPANY_COL]):
            raise SystemExit(f"{d['sheet']}!{d['row']} 有公司名，不能用本命令清空")
        before = copy.deepcopy(row['cells'])
        row['cells'] = []
        cleared.append({'sheet': d['sheet'], 'row': d['row'], 'cells_before': before})
    for s in master['sheets']:
        s['rows'] = [r for r in s['rows'] if r['cells']]
    ledger.setdefault('data_repairs', []).append({'id': uuid.uuid4().hex, 'action': 'clear_orphan_rows', 'rows': cleared,
                                                 'user_quote': quote.strip(), 'confirmed_at': now()})
    refresh_digest(master)
    save_master(ctx, master)
    save_json(ctx.ledger_path, ledger)
    return {'cleared_rows': [f"{c['sheet']}!{c['row']}" for c in cleared], 'excel_written': False}


def pending(ctx):
    master = ctx.load_master()
    refresh_digest(master)
    w = master.get('excel_written') or {}
    excel_sha = file_sha(ctx.book)
    if w:
        unwritten = w.get('data_sha256') != master['data_sha256'] or w.get('workbook_sha256') != excel_sha
    else:
        unwritten = diff_counts(ctx, master)[0] > 0
    ledger = ctx.ledger()
    deferred = any(d.get('data_sha256') == master['data_sha256'] for d in ledger.get('deferrals', []))
    waiting = sum(1 for e in ledger['entries'] if not e.get('applied_at') and not e.get('superseded_at'))
    return {'unwritten': unwritten, 'deferred_by_user': deferred, 'confirmed_status_changes_waiting': waiting,
            'excel_modified_outside_script': bool(w) and w.get('workbook_sha256') != excel_sha}


def defer(ctx, quote):
    if not quote or not quote.strip():
        raise SystemExit('必须提供用户“下一轮再写入”的原话')
    master = ctx.load_master()
    refresh_digest(master)
    ledger = ctx.ledger()
    ledger.setdefault('deferrals', []).append({'data_sha256': master['data_sha256'], 'user_quote': quote.strip(), 'deferred_at': now()})
    save_json(ctx.ledger_path, ledger)
    return {'deferred': True}


def query_rows(ctx, company=None, sheet=None, offset=0, limit=30):
    """只读：按公司或工作表分页读取主数据（AI 只读 JSON，不读 Excel）。"""
    if not company and not sheet:
        raise SystemExit('指定公司或工作表后分页读取')
    if offset < 0 or not 1 <= limit <= 200:
        raise SystemExit('分页范围无效：offset>=0，limit 1-200')
    master = ctx.load_master()
    out = []
    for s in master['sheets']:
        if sheet and s['name'] != sheet:
            continue
        names = {col: name for name, col in headers(s).items()} if any(r['row'] == 1 for r in s['rows']) else {}
        for r in s['rows']:
            if r['row'] == 1:
                continue
            values = {names.get(c['column'], str(c['column'])): c['value'] for c in r['cells']}
            if company and values.get('公司') != company:
                continue
            out.append({'sheet': s['name'], 'row': r['row'], 'values': values})
    return {'total': len(out), 'offset': offset, 'rows': out[offset:offset + limit]}


TRAINEE_PLAN = re.compile(r'管培|管理培训生|产培|产品培训生|Management Trainee|Graduate Program', re.I)


def _plain(v):
    return str(v.get('value', ''))[:10] if isinstance(v, dict) and v.get('type') == 'datetime' else v


def company_report(ctx, company):
    """只读：从主数据汇总一家公司的汇报口径（主链接、计划、管培/产培、岗位按状态、名额与占用）。"""
    master = ctx.load_master()
    counts = set((ctx.cfg.get('quotaRule') or {}).get('countsStatuses', []))

    def rows_of(name):
        sh = next((s for s in master['sheets'] if s['name'] == name), None)
        if not sh or not any(r['row'] == 1 for r in sh['rows']):
            return []
        idx = headers(sh)
        return [{h: _plain(cell_value(r, c)) for h, c in idx.items()} for r in sh['rows']
                if r['row'] > 1 and cell_value(r, idx.get(COMPANY_COL, 0)) == company]
    roles, placeholder, regions = [], None, []
    for name in ctx.sheets:
        s = sheet_of(master, name)
        idx = headers(s)
        for r in s['rows']:
            if r['row'] == 1 or cell_value(r, idx[COMPANY_COL]) != company:
                continue
            regions.append(name)
            v = {h: cell_value(r, c) for h, c in idx.items()}
            if row_kind(r, idx) == 'company':
                placeholder = v.get(ctx.column)
            else:
                roles.append({k: v.get(k) for k in ('岗位名称', '工作地点', '所属BU', '校招计划', ctx.column, '投递截止', '岗位链接')})
    if not regions:
        raise SystemExit('主数据中没有这家公司: ' + company)
    plans = rows_of('校招计划')
    by_status = {}
    for r in roles:
        by_status[r[ctx.column] or '（空）'] = by_status.get(r[ctx.column] or '（空）', 0) + 1
    quotas = []
    for _, v in quota_records(master):
        if v.get('公司') != company:
            continue
        lim, cur = v.get('投递上限'), v.get('当前投递数量')
        quotas.append({**{k: v.get(k) for k in ('范围', '适用计划/批次', '投递上限', '当前投递数量', '是否分BU可平行投递', '是否可以滚动投递')},
                       '剩余': lim - (cur or 0) if isinstance(lim, int) else '公司未披露'})
    return {'公司': company, '地区表': sorted(set(regions)), '公司行状态': placeholder,
            '校招主链接': [{k: v.get(k) for k in ('校招主链接', '适用届别', '投递窗口')} for v in rows_of('校招主链接')],
            '校招计划': {'数量': len([v for v in plans if v.get('计划')]),
                       '计划': [{k: v.get(k) for k in ('计划', '投递关系', '盘点状态', '核验日期')} for v in plans if v.get('计划')],
                       '管培/产培计划': [v.get('计划') for v in plans if v.get('计划') and TRAINEE_PLAN.search(str(v.get('计划')))]},
            '岗位': {'总数': len(roles), '按投递状态': by_status, '用户已选（计入名额）': sum(r[ctx.column] in counts for r in roles),
                   '列表': roles},
            '投递名额': quotas}


def main():
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--workspace', default=os.environ.get('CAMPUS_RECRUITING_WORKSPACE', '.'))
    ap.add_argument('--labels', default=os.environ.get('CAMPUS_STATUS_LABELS', '产品管理/投递状态标签.json'))
    sub = ap.add_subparsers(dest='cmd', required=True)
    sub.add_parser('check')
    sub.add_parser('pending')
    p = sub.add_parser('diff')
    p.add_argument('--limit', type=int, default=30, help='最多展开行数，0 为全部')
    sub.add_parser('write')
    p = sub.add_parser('query')
    p.add_argument('--company')
    p.add_argument('--sheet')
    p.add_argument('--offset', type=int, default=0)
    p.add_argument('--limit', type=int, default=30)
    p = sub.add_parser('report')
    p.add_argument('--company', required=True, action='append', help='公司名，可多次')
    for name in ('propose', 'confirm'):
        p = sub.add_parser(name)
        p.add_argument('--input', required=True, help='UTF-8 JSON：[{sheet,company,link|title|placeholder,to}]')
        if name == 'confirm':
            p.add_argument('--quote', required=True, help='用户确认原话')
            p.add_argument('--ref', default='', help='确认依据引用（对话/文件）')
    p = sub.add_parser('status')
    p.add_argument('--sheet', required=True)
    p.add_argument('--company', required=True)
    p.add_argument('--link')
    p.add_argument('--title')
    p = sub.add_parser('upsert')
    p.add_argument('--input', required=True, help='UTF-8 JSON：[{key:{sheet,company,link|title},values:{列:值}}]')
    p = sub.add_parser('reset-rows')
    p.add_argument('--input', required=True, help='UTF-8 JSON：[{sheet,company,link|title}] 要删除的岗位行')
    p.add_argument('--quote', default='', help='用户确认原话（--preview 时可省）')
    p.add_argument('--preview', action='store_true')
    p.add_argument('--drop-company', action='append', default=[], help='岗位删光后连同公司行整家删除的公司（可多次）')
    p.add_argument('--mark-research', action='store_true', help='公司仍有其它岗位时也把公司行改回“待搜索”，标明需重新搜索')
    p = sub.add_parser('link-exception')
    p.add_argument('--input', required=True, help='UTF-8 JSON：[{sheet,company,link,title,定位说明}]，撤销时给 "撤销": true')
    p.add_argument('--quote', default='', help='用户确认原话（--preview 时可省）')
    p.add_argument('--preview', action='store_true')
    p = sub.add_parser('relink')
    p.add_argument('--input', required=True, help='UTF-8 JSON：[{sheet,company,link,title,new_link}]，link/title 为该行当前值')
    p.add_argument('--quote', default='', help='用户确认原话（--preview 时可省）')
    p.add_argument('--preview', action='store_true')
    p = sub.add_parser('set-quota-rules')
    p.add_argument('--input', required=True, help='UTF-8 JSON：[{公司,范围,适用计划/批次,投递上限,是否分BU可平行投递,是否可以滚动投递,规则原文,来源/核验日期,原文出处文件}]；按公司整体覆盖。规则原文须为来源原文逐字摘抄（无官方说法写“公司未披露”），来源/核验日期格式“来源：<公司>-<岗位或页面> 核验日期：YYYY-MM-DD”')
    p = sub.add_parser('set-sheet-rows')
    p.add_argument('--sheet', required=True)
    p.add_argument('--input', required=True, help='UTF-8 JSON 二维列表，首行为表头；整页覆盖')
    p = sub.add_parser('drop-column')
    p.add_argument('--sheet', required=True)
    p.add_argument('--name', required=True)
    p = sub.add_parser('add-column')
    p.add_argument('--sheet', required=True)
    p.add_argument('--name', required=True)
    p = sub.add_parser('clear-orphans')
    p.add_argument('--input', required=True, help='UTF-8 JSON：[{sheet,row}]，只能是公司列为空的行')
    p.add_argument('--quote', required=True)
    p = sub.add_parser('defer')
    p.add_argument('--quote', required=True)
    a = ap.parse_args()
    ctx = Ctx(a.workspace, a.labels)

    def read(f):
        return json.loads(Path(f).read_text(encoding='utf-8'))
    code = 0
    if a.cmd == 'check':
        m = ctx.load_master()
        probs = validate_labels(ctx, m)
        fprobs = validate_fields(ctx, m) + validate_quota(m) + validate_link_exceptions(ctx, m)
        changed, extra = diff_counts(ctx, m)
        ch = status_changes(ctx, m)
        miss = unconfirmed(ctx, ch)
        leftover, over = quota_closure(ctx, m)
        legacy = [f"{QUOTA_SHEET}!{r['row']} {v.get('公司')}｜{v.get('范围')}" for r, v in quota_records(m)
                  if quota_source_problem(v.get('公司'), v.get('来源/核验日期'))]
        out = {'quota_full_but_unclosed': [p['text'] for p in leftover], 'quota_exceeded': over,
               'quota_rows_not_in_current_format': len(legacy), 'quota_rows_not_in_current_format_samples': legacy[:20],
               'label_problems': len(probs), 'label_problem_samples': probs[:20],
               'field_problems': len(fprobs), 'field_problem_samples': fprobs[:20],
               'cells_json_vs_excel_diff': changed, 'excel_only_sheets': extra,
               'status_changes': len(ch), 'status_changes_without_user_confirmation': len(miss),
               'link_changes_waiting_for_excel': len(identity_map(ctx))}
        code = 1 if (probs or fprobs or miss or leftover or over) else 0
    elif a.cmd == 'status':
        m = ctx.load_master()
        key = normalize_key({'sheet': a.sheet, 'company': a.company, 'link': a.link, 'title': a.title})
        s, idx, row = find_row(ctx, m, key)
        out = {'key': key, 'status': cell_value(row, idx[ctx.column]) or ''}
    elif a.cmd == 'propose':
        m = ctx.load_master()
        plan = propose(ctx, read(a.input), m)
        sim = copy.deepcopy(m)
        for p_ in plan:
            s_, idx_, row_ = find_row(ctx, sim, p_['key'])
            set_cell(row_, idx_[ctx.column], p_['to'])
        derived, over = quota_closure(ctx, sim)
        out = {'show_user_exactly': [p_['text'] for p_ in plan] + [d['text'] for d in derived],
               'user_changes': len(plan), 'quota_closure_changes': len(derived), 'quota_exceeded': over,
               'rule': '必须把 show_user_exactly 原样展示给用户并取得明确确认后，才能用同一 input 调用 confirm'}
    elif a.cmd == 'confirm':
        with exclusive(ctx):
            out = confirm(ctx, read(a.input), a.quote, a.ref)
    elif a.cmd == 'upsert':
        with exclusive(ctx):
            out = upsert(ctx, read(a.input))
    elif a.cmd == 'write':
        with exclusive(ctx):
            out = write(ctx)
    elif a.cmd == 'pending':
        out = pending(ctx)
    elif a.cmd == 'query':
        out = query_rows(ctx, a.company, a.sheet, a.offset, a.limit)
    elif a.cmd == 'report':
        out = [company_report(ctx, c) for c in a.company]
    elif a.cmd == 'diff':
        m = ctx.load_master()
        out = diff_details(ctx, m, a.limit or 10 ** 9)
    elif a.cmd == 'reset-rows':
        if a.preview:
            out = reset_rows(ctx, read(a.input), a.quote, preview=True, drop_companies=a.drop_company, mark_research=a.mark_research)
        else:
            with exclusive(ctx):
                out = reset_rows(ctx, read(a.input), a.quote, drop_companies=a.drop_company, mark_research=a.mark_research)
    elif a.cmd == 'link-exception':
        if a.preview:
            out = link_exception(ctx, read(a.input), a.quote, preview=True)
        else:
            with exclusive(ctx):
                out = link_exception(ctx, read(a.input), a.quote)
    elif a.cmd == 'relink':
        if a.preview:
            out = relink(ctx, read(a.input), a.quote, preview=True)
        else:
            with exclusive(ctx):
                out = relink(ctx, read(a.input), a.quote)
    elif a.cmd == 'set-quota-rules':
        with exclusive(ctx):
            out = set_quota_rules(ctx, read(a.input))
    elif a.cmd == 'set-sheet-rows':
        with exclusive(ctx):
            out = set_sheet_rows(ctx, a.sheet, read(a.input))
    elif a.cmd == 'drop-column':
        with exclusive(ctx):
            out = drop_column(ctx, a.sheet, a.name)
    elif a.cmd == 'add-column':
        with exclusive(ctx):
            out = add_column(ctx, a.sheet, a.name)
    elif a.cmd == 'clear-orphans':
        with exclusive(ctx):
            out = clear_orphans(ctx, read(a.input), a.quote)
    else:
        with exclusive(ctx):
            out = defer(ctx, a.quote)
    print(json.dumps(out, ensure_ascii=False, indent=2))
    sys.exit(code)


if __name__ == '__main__':
    main()
