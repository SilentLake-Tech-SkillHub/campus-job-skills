"""主数据固定编码：公司 C00001、站点 S0001、岗位 J000001。

规则（不可改）：
- 编码在实体第一次出现时由本模块分配一次，写进主表，永不修改、删除后不复用；
- 编码不含任何会变的信息；链接、名称都只是“值”，变了在原编码下覆盖；
- 新数据对应到已有实体只按固定规则做一次（公司：名称或已确认别名完全一致；站点：同一公司下域名完全一致），
  结果写成编码，之后只按编码查找，不再用值临时推算。
"""
import re
from urllib.parse import urlparse

COMPANY_SHEET = '公司主数据'
COMPANY_HEAD = ['公司编码', '公司名称', '别名', '状态']
SITE_SHEET = '投递站点'
SITE_HEAD = ['站点编码', '公司编码', '公司', '站点', '登录链接', '账号核对页', '状态']

COMPANY_CODE_COL, SITE_CODE_COL, ROLE_CODE_COL = '公司编码', '站点编码', '岗位编码'
APPLY_COL, ORDER_COL, CITY_COL = '投递链接', '志愿顺序', '城市选项'
# 岗位表新增列（按此顺序追加在表尾）
ROLE_SHEET_COLS = (COMPANY_CODE_COL, ROLE_CODE_COL, SITE_CODE_COL, APPLY_COL, ORDER_COL, CITY_COL)
# 其它带“公司”列的表只加公司编码
CODE_ONLY_SHEETS = ('校招主链接', '校招计划', '新加坡公司官网', '投递名额')
# 只能由固定命令写入的列（upsert 不能写）
LOCKED_COLS = (COMPANY_CODE_COL, ROLE_CODE_COL, SITE_CODE_COL, APPLY_COL, ORDER_COL)

WIDTH = {'C': 5, 'S': 4, 'J': 6}
PATTERN = {k: re.compile(r'^%s\d{%d}$' % (k, w)) for k, w in WIDTH.items()}
COUNTERS = '编码计数器'
ALIAS_SEP = '、'
SITE_ACTIVE, SITE_RETIRED = '在用', '停用'


def kind_of(code):
    for k, p in PATTERN.items():
        if isinstance(code, str) and p.match(code):
            return k
    return None


def origin(url):
    """站点 = 官网的协议 + 域名（+ 端口），照原样记录。官网只有 http 的（如民生银行 http://career.cmbc.com.cn:8080，用户 2026-10-27 选定）
    同样建站点；投递端仍只接受 https 链接，这类站点不能经 MCP 投递，组批时由链接规则列出原因。"""
    p = urlparse(str(url or ''))
    return f'{p.scheme}://{p.netloc}' if p.scheme in ('https', 'http') and p.netloc else None


# ---------- 工作表 ----------

def _sheet(master, name):
    return next((s for s in master['sheets'] if s['name'] == name), None)


def _head(sheet):
    first = next(r for r in sheet['rows'] if r['row'] == 1)
    return {c['value']: c['column'] for c in first['cells']}


def _cell(row, col):
    for c in row['cells']:
        if c['column'] == col:
            return c.get('value')
    return None


def _set(row, col, value):
    for c in row['cells']:
        if c['column'] == col:
            if value in (None, ''):
                row['cells'].remove(c)
            else:
                c['value'], c['type'] = value, 's' if isinstance(value, str) else 'n'
            return
    if value not in (None, ''):
        row['cells'].append({'column': col, 'value': value, 'type': 's' if isinstance(value, str) else 'n'})
        row['cells'].sort(key=lambda c: c['column'])


def _new_sheet(master, name, head):
    s = {'name': name, 'rows': [{'row': 1, 'cells': [{'column': i + 1, 'value': h, 'type': 's'} for i, h in enumerate(head)]}],
         'max_row': 1, 'max_column': len(head), 'state': 'visible', 'merged_ranges': [], 'tables': []}
    master['sheets'].append(s)
    return s


def add_columns(sheet, names):
    """在表尾追加缺少的列（已有的不动）。"""
    head = _head(sheet)
    first = next(r for r in sheet['rows'] if r['row'] == 1)
    for name in names:
        if name not in head:
            col = max(head.values()) + 1
            first['cells'].append({'column': col, 'value': name, 'type': 's'})
            head[name] = col
            sheet['max_column'] = max(sheet.get('max_column', 0), col)
    return head


def ensure_structure(master, status_sheets):
    """主表编码结构：两张主数据表 + 岗位表编码列 + 其它表公司编码列 + 计数器。可重复调用。"""
    if not _sheet(master, COMPANY_SHEET):
        _new_sheet(master, COMPANY_SHEET, COMPANY_HEAD)
    if not _sheet(master, SITE_SHEET):
        _new_sheet(master, SITE_SHEET, SITE_HEAD)
    for name in status_sheets:
        add_columns(_sheet(master, name), ROLE_SHEET_COLS)
    for name in CODE_ONLY_SHEETS:
        s = _sheet(master, name)
        if s and any(r['row'] == 1 for r in s['rows']) and '公司' in _head(s):
            add_columns(s, (COMPANY_CODE_COL,))
    master.setdefault(COUNTERS, {'C': 0, 'S': 0, 'J': 0})


def next_code(master, kind):
    counters = master.setdefault(COUNTERS, {'C': 0, 'S': 0, 'J': 0})
    counters[kind] = int(counters.get(kind) or 0) + 1
    return f'{kind}{counters[kind]:0{WIDTH[kind]}d}'


def _rows(master, name):
    s = _sheet(master, name)
    if not s:
        return None, {}, []
    head = _head(s)
    return s, head, [r for r in s['rows'] if r['row'] > 1]


def _append_row(sheet, values, head):
    row = {'row': max(r['row'] for r in sheet['rows']) + 1, 'cells': []}
    for k, v in values.items():
        _set(row, head[k], v)
    sheet['rows'].append(row)
    sheet['max_row'] = row['row']
    return row


# ---------- 公司 ----------

def companies(master):
    """公司编码 → {名称, 别名列表, 状态}。"""
    _, head, rows = _rows(master, COMPANY_SHEET)
    out = {}
    for r in rows:
        code = _cell(r, head['公司编码'])
        if code:
            aliases = [a.strip() for a in str(_cell(r, head['别名']) or '').split(ALIAS_SEP) if a.strip()]
            out[code] = {'名称': _cell(r, head['公司名称']), '别名': aliases, '状态': _cell(r, head['状态'])}
    return out


def resolve_company(master, name):
    """固定规则：公司名称或已确认别名完全一致。找不到返回 None；对上多家报错。"""
    name = str(name or '').strip()
    hits = [c for c, v in companies(master).items() if name and (v['名称'] == name or name in v['别名'])]
    if len(hits) > 1:
        raise SystemExit(f'公司名「{name}」对上多个公司编码：{hits}，主数据需先修正')
    return hits[0] if hits else None


def ensure_company(master, name):
    code = resolve_company(master, name)
    if code:
        return code
    name = str(name or '').strip()
    if not name:
        raise SystemExit('公司名为空，不能建公司编码')
    s, head, _ = _rows(master, COMPANY_SHEET)
    code = next_code(master, 'C')
    _append_row(s, {'公司编码': code, '公司名称': name, '状态': '在用'}, head)
    return code


def company_name(master, code):
    v = companies(master).get(code)
    return v['名称'] if v else None


# ---------- 站点 ----------

def sites(master):
    """站点编码 → {公司编码, 站点, 登录链接, 账号核对页, 状态, row}。"""
    _, head, rows = _rows(master, SITE_SHEET)
    out = {}
    for r in rows:
        code = _cell(r, head['站点编码'])
        if code:
            out[code] = {h: _cell(r, head[h]) for h in SITE_HEAD if h != '站点编码'}
            out[code]['row'] = r
    return out


def site_of(master, company_code, site_origin):
    """固定规则：同一公司编码下，在用站点的域名完全一致。"""
    hits = [c for c, v in sites(master).items()
            if v['公司编码'] == company_code and v['站点'] == site_origin and v['状态'] != SITE_RETIRED]
    if len(hits) > 1:
        raise SystemExit(f'{company_code} 的站点 {site_origin} 对上多个站点编码：{hits}')
    return hits[0] if hits else None


def ensure_site(master, company_code, site_origin):
    if not site_origin:
        raise SystemExit('岗位链接不是 https 官网链接，不能确定站点')
    code = site_of(master, company_code, site_origin)
    if code:
        return code
    s, head, _ = _rows(master, SITE_SHEET)
    code = next_code(master, 'S')
    _append_row(s, {'站点编码': code, '公司编码': company_code, '公司': company_name(master, company_code),
                    '站点': site_origin, '状态': SITE_ACTIVE}, head)
    return code


def set_site_value(master, site_code, field, value):
    s, head, _ = _rows(master, SITE_SHEET)
    row = sites(master)[site_code]['row']
    _set(row, head[field], value)


# ---------- 公司名显示列 ----------

def sync_company_names(master):
    """各表“公司”列只是公司主数据名称的显示：保存前一律按公司编码重写。"""
    names = {c: v['名称'] for c, v in companies(master).items()}
    for s in master['sheets']:
        if s['name'] == COMPANY_SHEET or not any(r['row'] == 1 for r in s['rows']):
            continue
        head = _head(s)
        if COMPANY_CODE_COL not in head or '公司' not in head:
            continue
        for r in s['rows']:
            if r['row'] == 1:
                continue
            code = _cell(r, head[COMPANY_CODE_COL])
            if code in names and _cell(r, head['公司']) != names[code]:
                _set(r, head['公司'], names[code])


# ---------- 校验 ----------

def validate(master, status_sheets, link_col='岗位链接', name_col='岗位名称'):
    """编码门禁：格式、唯一、引用存在、站点与岗位链接一致、计数器不小于已用编码。"""
    probs = []
    counters = master.get(COUNTERS) or {}
    comp = companies(master)
    used = {'C': set(), 'S': set(), 'J': set()}
    names = {}
    for code, v in comp.items():
        if kind_of(code) != 'C':
            probs.append(f'{COMPANY_SHEET} 公司编码格式不对：{code}')
        if code in used['C']:
            probs.append(f'{COMPANY_SHEET} 公司编码重复：{code}')
        used['C'].add(code)
        for n in [v['名称']] + v['别名']:
            if not n:
                probs.append(f'{COMPANY_SHEET} {code} 公司名称为空')
            elif n in names and names[n] != code:
                probs.append(f'{COMPANY_SHEET} 名称/别名「{n}」同时属于 {names[n]} 和 {code}')
            names[n] = code
    site_map = sites(master)
    pairs = {}
    for code, v in site_map.items():
        if kind_of(code) != 'S':
            probs.append(f'{SITE_SHEET} 站点编码格式不对：{code}')
        if code in used['S']:
            probs.append(f'{SITE_SHEET} 站点编码重复：{code}')
        used['S'].add(code)
        if v['公司编码'] not in comp:
            probs.append(f'{SITE_SHEET} {code} 的公司编码不存在：{v["公司编码"]}')
        if origin(v['站点']) != v['站点']:
            probs.append(f'{SITE_SHEET} {code} 站点必须是官网的协议加域名（不带路径）：{v["站点"]}')
        if v['状态'] != SITE_RETIRED:
            k = (v['公司编码'], v['站点'])
            if k in pairs:
                probs.append(f'{SITE_SHEET} {code} 与 {pairs[k]} 是同一公司同一站点')
            pairs[k] = code
        for f in ('登录链接', '账号核对页'):
            if v[f] and origin(v[f]) != v['站点']:
                probs.append(f'{SITE_SHEET} {code} {f}不在本站点：{v[f]}')
    for name in status_sheets:
        s = _sheet(master, name)
        head = _head(s)
        for col in ROLE_SHEET_COLS:
            if col not in head:
                probs.append(f'{name} 缺少列「{col}」')
        if any(col not in head for col in ROLE_SHEET_COLS):
            continue
        for r in s['rows']:
            if r['row'] == 1 or not any(c.get('value') not in (None, '') for c in r['cells']):
                continue
            where = f'{name}!{r["row"]}'
            ccode, jcode, scode = (_cell(r, head[c]) for c in (COMPANY_CODE_COL, ROLE_CODE_COL, SITE_CODE_COL))
            is_role = bool(_cell(r, head.get(name_col, 0)) or _cell(r, head.get(link_col, 0)))
            if ccode not in comp:
                probs.append(f'{where} 公司编码缺失或不存在：{ccode}')
            if not is_role:
                if jcode or scode or _cell(r, head[APPLY_COL]):
                    probs.append(f'{where} 公司行不能有岗位编码/站点编码/投递链接')
                continue
            if kind_of(jcode) != 'J':
                probs.append(f'{where} 岗位编码缺失或格式不对：{jcode}')
            elif jcode in used['J']:
                probs.append(f'{where} 岗位编码重复：{jcode}')
            used['J'].add(jcode)
            link = _cell(r, head[link_col]) if link_col in head else None
            site = site_map.get(scode)
            if link and not site:
                probs.append(f'{where} 站点编码缺失或不存在：{scode}')
            if site:
                if site['公司编码'] != ccode:
                    probs.append(f'{where} 站点 {scode} 不属于公司 {ccode}')
                if link and origin(link) != site['站点']:
                    probs.append(f'{where} 岗位链接域名与站点 {scode}（{site["站点"]}）不一致')
                apply = _cell(r, head[APPLY_COL])
                if apply and origin(apply) != site['站点']:
                    probs.append(f'{where} 投递链接不在站点 {scode}（{site["站点"]}）')
    for name in CODE_ONLY_SHEETS:
        s = _sheet(master, name)
        if not s or not any(r['row'] == 1 for r in s['rows']):
            continue
        head = _head(s)
        if '公司' not in head:
            continue
        if COMPANY_CODE_COL not in head:
            probs.append(f'{name} 缺少列「{COMPANY_CODE_COL}」')
            continue
        for r in s['rows']:
            if r['row'] > 1 and any(c.get('value') not in (None, '') for c in r['cells']):
                if _cell(r, head[COMPANY_CODE_COL]) not in comp:
                    probs.append(f'{name}!{r["row"]} 公司编码缺失或不存在')
    for k, codes in used.items():
        top = max((int(c[1:]) for c in codes if kind_of(c) == k), default=0)
        if int(counters.get(k) or 0) < top:
            probs.append(f'{COUNTERS}.{k}={counters.get(k)} 小于已用的最大编号 {top}')
    return probs


def assign_missing_codes(master, status_sheets, link_col='岗位链接', name_col='岗位名称'):
    """给还没有编码的行分配编码（第一次出现）：公司按名称/别名完全一致对应或新建；岗位新建；
    站点按“同一公司 + 岗位链接域名完全一致”对应或新建。已有编码的行不动。返回新分配数量。"""
    ensure_structure(master, status_sheets)
    made = {'C': 0, 'S': 0, 'J': 0}

    def company(name):
        before = master[COUNTERS]['C']
        code = ensure_company(master, name)
        made['C'] += master[COUNTERS]['C'] - before
        return code
    for name in status_sheets:
        s = _sheet(master, name)
        head = _head(s)
        for r in sorted(s['rows'], key=lambda r: r['row']):
            if r['row'] == 1 or not any(c.get('value') not in (None, '') for c in r['cells']):
                continue
            cname = _cell(r, head['公司'])
            if not cname:
                continue
            ccode = _cell(r, head[COMPANY_CODE_COL]) or company(cname)
            _set(r, head[COMPANY_CODE_COL], ccode)
            if not (_cell(r, head.get(name_col, 0)) or _cell(r, head.get(link_col, 0))):
                continue
            if not _cell(r, head[ROLE_CODE_COL]):
                _set(r, head[ROLE_CODE_COL], next_code(master, 'J'))
                made['J'] += 1
            site_origin = origin(_cell(r, head[link_col]))
            if site_origin and not _cell(r, head[SITE_CODE_COL]):
                before = master[COUNTERS]['S']
                _set(r, head[SITE_CODE_COL], ensure_site(master, ccode, site_origin))
                made['S'] += master[COUNTERS]['S'] - before
    for name in CODE_ONLY_SHEETS:
        s = _sheet(master, name)
        if not s or not any(r['row'] == 1 for r in s['rows']):
            continue
        head = _head(s)
        if '公司' not in head:
            continue
        for r in s['rows']:
            if r['row'] > 1 and _cell(r, head['公司']) and not _cell(r, head[COMPANY_CODE_COL]):
                _set(r, head[COMPANY_CODE_COL], company(_cell(r, head['公司'])))
    return made
