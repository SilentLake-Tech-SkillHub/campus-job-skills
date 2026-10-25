#!/usr/bin/env python3
"""投递状态与 JSON→Excel Hook：写入保护全局生效，同步提醒可按对话关闭。

用法：status_sync_hook.py --workspace <工作区根> --host claude|codex|agy [--event <宿主事件名>]
stdin 为宿主 Hook payload。

- 会话/每轮开始：注入标签体系与门禁规则；若 JSON 有未写入 Excel 的改动且用户未选择延后，要求先询问用户。
- 工具调用前：拦截绕过固定脚本直接修改 Excel、主数据 JSON、确认台账、标签配置的写操作。
- 对话结束（Stop）：存在未写入且用户未选择延后的改动时阻止结束，要求询问“现在写入还是下一轮写入”。
"""
import argparse
import json
import os
import re
import shlex
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from host_payload import (call_parts, _session, PROMPT_KEYS, _prompt, QUESTION_REPLY_TAG, QUESTION_REPLY,
                          QUESTION_CALL_ID, _approval_prompts)

SCRIPT = Path(__file__).with_name('json_excel_sync.py')
EDIT_TOOL = re.compile(r'(write|edit|patch|replace|create|insert|notebook)', re.I)
SHELL_TOOL = re.compile(r'(bash|shell|exec|command|terminal|run)', re.I)
SHELL_WRITE = re.compile(r'(\.save\(|to_excel|write_text|write_bytes|open\([^)]*["\'][wa]b?["\']|>\s*\S|\bcp\b|\bmv\b|sed\s+-i|tee\b|truncate|json\.dump\()', re.I)


def emit(obj):
    sys.stdout.write(json.dumps(obj, ensure_ascii=False))


LABELS_FILE = '产品管理/投递状态标签.json'


def sync_reminders_enabled(root, payload):
    """按真实会话身份关闭提醒与 Stop 同步检查，保留 PreToolUse 写入保护。"""
    sid = str(payload.get('session_id') or payload.get('sessionId')
              or payload.get('conversation_id') or payload.get('thread_id')
              or os.environ.get('DSH_SESSION_ID') or os.environ.get('CODEX_THREAD_ID')
              or os.environ.get('CODEX_SESSION_ID') or '')
    if not sid:
        return True
    try:
        scope = json.loads((root / '.codex/hooks/scope.json').read_text(encoding='utf-8'))
        session = scope.get('sessions', {}).get(sid, {})
    except (OSError, ValueError, AttributeError):
        return True
    return not isinstance(session, dict) or session.get('status_sync_reminders') is not False


def load_cfg(root):
    try:
        return json.loads((root / LABELS_FILE).read_text(encoding='utf-8'))
    except (OSError, ValueError):
        return None


def label_groups(root, cfg):
    """与 json_excel_sync 同一个加载入口读取全部标签组（一组一个文件）。"""
    sys.path.insert(0, str(SCRIPT.parent))
    import json_excel_sync
    try:
        return json_excel_sync.load_label_groups(root, cfg, root / LABELS_FILE)[1]
    except SystemExit as e:
        return [{'path': LABELS_FILE, 'name': '标签组加载失败', 'rowKind': 'role', 'labels': [], 'error': str(e)}]


def protected(cfg):
    batch = (cfg.get('companyStore') or {}).get('batchFile', '岗位交接/投递批次.json')
    # 投递批次只由 company_store.py 写；投递回传只由投递组件的 application_closeout 写
    return [cfg['workbook'], cfg['masterData'], cfg['confirmationLedger'], LABELS_FILE, *cfg.get('labelGroups', []),
            batch, str(Path(batch).with_name('投递回传.json')), '产品管理/user_applicaiton_requirements_temp.json',
            '.runtime/user_application_requirements_state.json']


def sync(root, *args):
    p = subprocess.run([sys.executable, '-B', str(SCRIPT), '--workspace', str(root), *args],
                       capture_output=True, text=True, encoding='utf-8', timeout=25)
    try:
        return json.loads(p.stdout)
    except ValueError:
        return {'error': (p.stderr or p.stdout).strip()[:300]}


CONFIG_FILE = '产品管理/求职配置.md'


def browser_rules(root):
    """浏览器规则只有一个位置：求职配置.md 的「## 浏览器」一节，每轮原样注入，不在别处复述。"""
    try:
        text = (Path(root) / CONFIG_FILE).read_text(encoding='utf-8')
    except OSError:
        return f'\n【浏览器规则】读取 {CONFIG_FILE} 失败，开页前先读该文件「浏览器」一节。'
    m = re.search(r'^## 浏览器[^\n]*\n(.*?)(?=^## |\Z)', text, re.M | re.S)
    if not m or not m.group(1).strip():
        return f'\n【浏览器规则】{CONFIG_FILE} 缺少「## 浏览器」一节，开页前先请用户补充。'
    return f'\n【浏览器规则（唯一来源：{CONFIG_FILE}「浏览器」，开页、切换或降级前必须遵守）】\n' + m.group(1).strip()


def rules_text(cfg, root=None):
    kinds = {'role': '岗位行', 'company': '公司行'}
    groups = label_groups(Path(root), cfg) if root is not None else [
        {'path': LABELS_FILE, 'rowKind': cfg.get('rowKind', 'role'), 'labels': cfg.get('labels', [])}]
    lines = []
    for g in groups:
        if g.get('error'):
            lines.append('  ' + g['error'])
            continue
        rel = (os.path.relpath(g['path'], root) if root is not None and os.path.isabs(g['path']) else str(g['path'])).replace(os.sep, '/')
        lines.append(f"  {kinds.get(g['rowKind'], g['rowKind'])}只能用（{rel}）："
                     + '；'.join(f"「{l['value']}」={l.get('definition', '')}" for l in g['labels']))
    return ('【投递状态门禁（所有 Agent 必须遵守）】\n'
            '1. 投递状态只能是标签组文件中的标签（一组一个文件：岗位组 产品管理/投递状态标签.json，其它组见其 labelGroups），每行有且只有一个、不得为空、不得追加说明；标签只能由固定脚本写入，写入前由脚本统一检查：\n'
            + '\n'.join(lines) + '\n'
            + '   公司行=岗位名称与岗位链接均为空的行；每家公司在其出现的每张地区表固定一行公司行，有没有岗位都保留，新增岗位不占用公司行。\n'
            + '2. 主数据是 岗位交接/工作簿全量.json，与 Excel 完全对照。Excel 只给人看：AI 禁止读取或修改任何 .xlsx，只读 JSON；Excel 只能由 json_excel_sync.py write 从 JSON 写出，禁止绕过脚本改 JSON。\n'
            '3. 任何状态变更：先 propose，把输出的“当前状态是X，更新状态是Y”原样给用户看，取得用户明确确认后才能 confirm（附用户原话）。已有岗位换链接只能用 relink（先 --preview 给用户看“岗位链接 X → Y”，确认后带原话执行），不能用 upsert 直接改。官网无独立详情页的岗位，链接例外备注只能用 link-exception 写（先 --preview 给用户看，确认后执行）。\n'
            '4. 名额收口：用户的选择占满本轮名额后，同范围剩余“岗位待筛选”由 confirm 自动改为“岗位不符合用户要求”，并在 propose 中一并展示；名额未被用户选满前不得剔除任何岗位。\n'
            '5. 每次对话结束前，若 JSON 有未写入 Excel 的改动，必须问用户：现在写入 Excel，还是下一轮再写入；用户选现在写入则 write，选下一轮则 defer（附原话）。\n'
            '6. 记录唯一：每类信息只有一个存放位置、只保留最新一版，更新=覆盖原值；禁止追加“旧结论/历史研究记录”，禁止另存副本/历史/备份，禁止按批次另起同类文件；回滚靠 git。岗位交接/ 下只允许固定文件（见 AGENTS.md）。\n'
            '7. 组件软锁：.agents/skills/campus-job-application/、.agents/skills/job-application-assistant-v2/ 的代码和 Hook 配置默认只读；用户在对话里明确批准修改组件后，本会话可以修改（12 小时内有效，用户说“收回授权”即失效），或由用户本人把 产品管理/组件维护.json 的“开启”改为 true。需要修改时先说明要改什么并请用户批准；AI 不得写开关文件和批准记录（.runtime/组件维护批准.json）。')


def pending_details(root, limit=30):
    """JSON 写入 Excel 的实际逐格差异（由固定脚本读取，AI 本身不读 Excel）。"""
    d = sync(root, 'diff', '--limit', str(limit))
    return d.get('text') or ('差异明细生成失败：' + str(d.get('error', '未知错误')))


def pending_text(root):
    p = sync(root, 'pending')
    if p.get('unwritten') and not p.get('deferred_by_user'):
        extra = '（另检测到 Excel 被脚本以外修改，写入会以 JSON 覆盖）' if p.get('excel_modified_outside_script') else ''
        return (f"\n【待处理】JSON 有未写入 Excel 的改动（已确认待写入的状态变更 {p.get('confirmed_status_changes_waiting', 0)} 行）{extra}。"
                '请先把下面的实际差异明细原样展示给用户（禁止只报总数），再询问用户：现在写入 Excel，还是下一轮再写入。\n'
                + pending_details(root)), p
    return '', p


HANDOFF_TOOLS = ('company_batch_create', 'role_queue_import')


def handoff_violation(root, payload):
    """正向门禁：投递组件导入批次前，投递批次必须已由用户评审确认且确认后未被改动。"""
    call = payload.get('toolCall') or {}
    tool = str(payload.get('tool_name') or payload.get('toolName') or call.get('name') or '')
    raw = json.dumps(payload.get('tool_input') or payload.get('arguments') or call.get('arguments') or {}, ensure_ascii=False)
    if not any(tool.endswith(t) or t in raw and 'ToolName' in raw for t in HANDOFF_TOOLS):
        return None
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import company_store
    st = company_store.Store(root, '产品管理/投递状态标签.json')
    problem = company_store.verify_handoff(st.batch())
    return ('正向门禁：' + problem + '。请先运行 company_store.py forward 生成评审页，用户确认后 forward-confirm。') if problem else None


EXCEL = re.compile(r'\.xls[xm]?\b', re.I)
PATCH_TARGET = re.compile(r'^\*\*\* (?:(?:Update|Add|Delete) File|Move to): *(.+?)\s*$', re.M)
PATH_KEYS = ('file_path', 'path', 'notebook_path', 'filePath', 'TargetFile', 'AbsolutePath')


def write_targets(ti):
    """编辑类调用真正改写的文件：apply_patch 取补丁头里的目标路径，Write/Edit 取路径参数；无法确定时返回 None。"""
    texts = [ti] if isinstance(ti, str) else [v for v in ti.values() if isinstance(v, str)] if isinstance(ti, dict) else []
    for t in texts:
        if '*** Begin Patch' in t:
            return PATCH_TARGET.findall(t)
    if isinstance(ti, dict):
        found = [str(ti[k]) for k in PATH_KEYS if ti.get(k)]
        if found:
            return found
    return None


def excel_violation(payload):
    """Excel 只给人看：任何 Agent 不得以任何方式读取或改写 .xlsx，唯一放行固定写入脚本。"""
    tool, ti = call_parts(payload)
    raw = json.dumps(ti, ensure_ascii=False)
    targets = write_targets(ti) if EDIT_TOOL.search(tool) else None
    subject = ' '.join(targets) if targets is not None else raw
    if EXCEL.search(subject) and 'json_excel_sync.py' not in subject:
        return ('Excel 门禁：Excel 只给人看，AI 禁止读取或修改任何 .xlsx。请改读 岗位交接/工作簿全量.json'
                '（与 Excel 完全对照的主数据）；写 Excel 只能用 json_excel_sync.py write。')
    return None


HANDOFF_ALLOWED = re.compile(r'^岗位交接/(工作簿全量\.json|投递批次\.json|正向评审\.html|投递回传\.json|反向评审\.html|公司/[^/]+/[^/]+/(公司|登录(@[^/]+)?|表单定义(@[^/]+)?)\.json)$')
STACK_NAME = re.compile(r'(^|/)(历史|备份|快照|存档)(/|$)|上一版|副本|[-_.]v\d+\.[A-Za-z]+$|\.(bak|orig)$')
PATH_TOKEN = re.compile(r'[^\s\'"<>|;&()]*(?:岗位交接/|历史/|备份|快照|上一版|副本|\.bak)[^\s\'"<>|;&()]*')


def stack_violation(root, payload):
    """记录唯一门禁：同一信息只保留最新一版、覆盖更新；不得在工作区另存副本/历史/备份，岗位交接/ 只允许固定文件。"""
    root = Path(root).resolve()
    tool, ti = call_parts(payload)
    if EDIT_TOOL.search(tool) and not (isinstance(ti, dict) and 'command' in ti):
        targets = write_targets(ti) or []
    elif SHELL_TOOL.search(tool) or isinstance(ti, dict) and 'command' in ti:
        raw = ti.get('command') if isinstance(ti, dict) and isinstance(ti.get('command'), str) else json.dumps(ti, ensure_ascii=False)
        if not SHELL_WRITE.search(raw):
            return None
        targets = [t for t in PATH_TOKEN.findall(raw) if ':' not in t.split('岗位交接/')[0]]
    else:
        return None
    bad = []
    for t in targets:
        t = str(t).strip().strip('\'"')
        if Path(t).is_absolute() or t.startswith('/'):
            try:
                r = Path(t).resolve().relative_to(root).as_posix()
            except ValueError:
                continue
        else:
            r = t[2:] if t.startswith('./') else t
        if (r.startswith('岗位交接/') and not HANDOFF_ALLOWED.match(r)) or STACK_NAME.search(r):
            bad.append(r)
    if not bad:
        return None
    return ('记录唯一门禁：同一信息只保留最新一版、覆盖原文件更新，禁止另存副本/历史/备份/按批次另起同类文件（回滚靠 git）。'
            '岗位交接/ 下只允许 工作簿全量.json、投递批次.json、正向评审.html、投递回传.json、反向评审.html、公司/<地区>/<公司>/{公司.json,登录[@域名].json,表单定义[@域名].json}；'
            '过程文件放工作区以外的临时目录，用完即删。被拦截：' + '、'.join(bad[:5]))


def tool_violation(cfg, payload, root=None):
    tool, ti = call_parts(payload)
    raw = json.dumps(ti, ensure_ascii=False)
    files = protected(cfg)
    targets = write_targets(ti) if EDIT_TOOL.search(tool) else None
    if targets is not None:
        # 按完整路径判断：相对路径以本次调用的 cwd 解析后与工作区内的受保护文件比较，
        # 其他目录里的同名文件不受影响。
        base = Path(payload.get('cwd') or root or '.')

        def hit(t, f):
            t = str(t).strip().strip('\'"')
            if root is not None:
                p = Path(t) if Path(t).is_absolute() or t.startswith('/') else base / t
                try:
                    return p.resolve() == (Path(root).resolve() / f).resolve()
                except OSError:
                    pass
            return t == f or t.replace('\\', '/').endswith('/' + f)
        return [f for f in files if any(hit(t, f) for t in targets)] or None
    if SHELL_TOOL.search(tool) and not EDIT_TOOL.search(tool) or isinstance(ti, dict) and 'command' in ti:
        hits = [f for f in files if f in raw or Path(f).name in raw]
        if hits and 'json_excel_sync.py' not in raw and SHELL_WRITE.search(raw):
            if 'application_requirements.py' in raw:
                hits = [f for f in hits if f not in ('产品管理/user_applicaiton_requirements_temp.json',
                                                    '.runtime/user_application_requirements_state.json')]
            return hits or None
        return None
    if EDIT_TOOL.search(tool):
        targets = write_targets(ti)
        if targets is None:
            hits = [f for f in files if f in raw or Path(f).name in raw]
        else:
            hits = [f for f in files if any(t == f or t.endswith('/' + f) or Path(t).name == Path(f).name for t in targets)]
        return hits or None
    return None


# 组件软锁：两个组件的代码与 Hook 配置默认锁定。两种放开方式：
#   1) 用户在对话里明确批准修改组件（如“批准修改组件”“授权你改 Hook”），本会话放开，用户说“收回授权/锁上组件”或 12 小时后失效；
#      批准只从用户本人发出的消息（UserPromptSubmit）识别，由 Hook 记录到 APPROVALS，AI 不能写这个记录。
#   2) 用户本人把维护开关文件的“开启”改为 true（长期放开）。开关文件与批准记录 AI 任何时候都不能写。
LOCKED = ('.agents/skills/campus-job-application/', '.agents/skills/job-application-assistant-v2/',
          '.claude/settings.json', '.codex/hooks.json')
SWITCH = '产品管理/组件维护.json'
APPROVALS = '.runtime/组件维护批准.json'
APPROVAL_HOURS = 12
COMPONENT_WORD = re.compile(r'组件|插件|[Hh]ook|钩子|[Ss]kill')
APPROVE_MODIFY = re.compile(r'(批准|授权|同意|允许).{0,6}(修改|改|维护|更新)')
CONDITIONAL = re.compile(r'情况下|的话|前提|时才|之后才|需要.{0,4}(批准|授权)|是否')
NEGATED = re.compile(r'(不|别|禁止|不要|不准|不许)\s*(批准|授权|同意|允许|修改|改|维护)')
REVOKE = re.compile(r'(收回|撤销|撤回|取消|结束)\s*(组件)?\s*(维护)?\s*(授权|批准|维护)|锁上组件|锁回组件|恢复只读')
SEGMENT = re.compile(r'&&|\|\||;|\n|\|')
REDIRECT = re.compile(r'(?:^|[^0-9&>])>{1,2}\s*([^\s&;|>][^\s;&|]*)')
IN_PLACE = {'sed', 'perl', 'truncate', 'touch', 'chmod', 'chown', 'rm', 'rmdir', 'unlink', 'patch', 'ln', 'install', 'tee', 'mkdir'}
COPY = {'cp', 'mv', 'rsync', 'ditto'}
GIT_MUTATE = {'checkout', 'restore', 'reset', 'apply', 'am', 'stash', 'clean', 'pull', 'merge', 'rebase', 'cherry-pick', 'rm', 'mv', 'switch', 'revert'}
PKG_MUTATE = {'install', 'i', 'ci', 'uninstall', 'remove', 'update', 'add', 'upgrade'}
INLINE_WRITE = re.compile(r'(write_text|write_bytes|open\([^)]*["\'][wax]b?\+?["\']|json\.dump\(|writeFile|appendFile|unlink|rmSync|shutil\.|os\.remove|os\.rename)')


def maintenance_on(root):
    try:
        return json.loads((root / SWITCH).read_text(encoding='utf-8')).get('开启') is True
    except (OSError, ValueError, AttributeError):
        return False


def _debug(root, event, payload, found):
    """只记字段名和是否取到原话，不记内容；排查宿主 payload 用，覆盖写。"""
    try:
        f = root / '.runtime/组件软锁诊断.json'
        f.parent.mkdir(parents=True, exist_ok=True)
        f.write_text(json.dumps({'event': event, 'keys': sorted(payload)[:40], 'session': _session(payload),
                                 'promptFound': found, 'at': __import__('time').strftime('%Y-%m-%d %H:%M:%S')},
                                ensure_ascii=False, indent=2), encoding='utf-8')
    except OSError:
        pass


def _load_approvals(root):
    try:
        d = json.loads((root / APPROVALS).read_text(encoding='utf-8'))
        return d if isinstance(d, dict) else {}
    except (OSError, ValueError):
        return {}


def record_approval(root, payload):
    """只看用户本人发出的消息：明确批准修改组件 → 记录本会话批准；收回 → 删除。返回给 AI 的提示。"""
    prompts = _approval_prompts(payload)
    _debug(root, 'UserPromptSubmit', payload, any(p.strip() for p in prompts))
    if not any(p.strip() for p in prompts):
        return ''
    sid, data = _session(payload), _load_approvals(root)
    now = __import__('time').time()
    data = {k: v for k, v in data.items() if v.get('expiresAt', 0) > now}
    note = ''
    approval_quotes = [p.strip() for p in prompts if COMPONENT_WORD.search(p) and APPROVE_MODIFY.search(p)
                       and not NEGATED.search(p) and not CONDITIONAL.search(p)]
    if any(REVOKE.search(p) for p in prompts):
        data.pop(sid, None)
        note = '\n【组件软锁】用户已收回本会话的组件修改批准，组件恢复只读。'
    elif approval_quotes:
        hours = 2 if sid == '_no_session' else APPROVAL_HOURS
        data[sid] = {'userQuote': '\n'.join(approval_quotes)[:300], 'approvedAt': now, 'expiresAt': now + hours * 3600}
        note = f'\n【组件软锁】已记录用户批准：本会话可以修改搜索组件、投递组件和 Hook 配置（{hours} 小时内有效，用户说“收回授权”即失效）。修改仍须按用户的具体要求进行，并逐项汇报。'
    else:
        return '\n【组件软锁】本会话已获用户批准修改组件。' if sid in data else ''
    f = root / APPROVALS
    f.parent.mkdir(parents=True, exist_ok=True)
    f.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    return note


def approved(root, payload):
    v = _load_approvals(root).get(_session(payload))
    return bool(v) and v.get('expiresAt', 0) > __import__('time').time()


def _rel(root, cwd, t):
    t = str(t).strip().strip('\'"')
    if not t or t.startswith('-'):
        return None
    p = Path(t).expanduser()
    p = p if p.is_absolute() else Path(cwd) / p
    try:
        return Path(os.path.normpath(str(p))).relative_to(root).as_posix()
    except ValueError:
        return None


def _locked(rel, include_switch=True):
    if rel is None:
        return False
    if include_switch and rel == SWITCH:
        return True
    return any(rel == l.rstrip('/') or (l.endswith('/') and rel.startswith(l)) for l in LOCKED)


def _shell_targets(root, cwd, cmd):
    """粗略解析 shell 命令里会被改写的路径（相对路径按 cd 后的目录解析）。"""
    out = []
    if '*** Begin Patch' in cmd:
        out += [_rel(root, cwd, t) for t in PATCH_TARGET.findall(cmd)]
    for seg in SEGMENT.split(cmd):
        out += [_rel(root, cwd, t) for t in REDIRECT.findall(seg) if t not in ('/dev/null',)]
        try:
            toks = shlex.split(seg)
        except ValueError:
            toks = seg.split()
        while toks and (re.match(r'^\w+=', toks[0]) or toks[0] in ('sudo', 'env', 'command', 'xargs')):
            toks = toks[1:]
        if not toks:
            continue
        name, args = Path(toks[0]).name, toks[1:]
        plain = [a for a in args if not a.startswith('-')]
        if name in ('cd', 'pushd') and plain:
            cwd = str(Path(cwd) / Path(plain[0]).expanduser()) if not Path(plain[0]).is_absolute() else plain[0]
        elif name in IN_PLACE:
            if name in ('sed', 'perl') and not any(a.startswith('-i') or a.startswith('-pi') for a in args):
                continue
            out += [_rel(root, cwd, a) for a in plain]
        elif name in COPY and plain:
            out.append(_rel(root, cwd, plain[-1]))
            if name == 'mv':
                out += [_rel(root, cwd, a) for a in plain[:-1]]
        elif name == 'git':
            d = cwd
            if '-C' in args and args.index('-C') + 1 < len(args):
                d = str(Path(cwd) / args[args.index('-C') + 1])
            sub = [a for a in args if not a.startswith('-') and a != (args[args.index('-C') + 1] if '-C' in args and args.index('-C') + 1 < len(args) else None)]
            if sub and sub[0] in GIT_MUTATE:
                out.append(_rel(root, d, '.'))
        elif name in ('npm', 'pnpm', 'yarn') and plain and plain[0] in PKG_MUTATE:
            out.append(_rel(root, cwd, '.'))
        elif name.startswith(('python', 'node', 'ruby', 'perl', 'osascript')) and INLINE_WRITE.search(seg):
            out += [_rel(root, cwd, t) for t in re.findall(r'[\w./\u4e00-\u9fff@-]+', seg) if '/' in t]
    return out


def component_violation(root, payload):
    tool, ti = call_parts(payload)
    cwd = payload.get('cwd') or str(root)
    if isinstance(ti, dict) and isinstance(ti.get('command'), str) or SHELL_TOOL.search(tool) and not EDIT_TOOL.search(tool):
        raw = ti.get('command') if isinstance(ti, dict) and isinstance(ti.get('command'), str) else json.dumps(ti, ensure_ascii=False)
        targets = _shell_targets(root, cwd, raw)
    elif EDIT_TOOL.search(tool):
        found = write_targets(ti)
        if found is None:
            raw = json.dumps(ti, ensure_ascii=False)
            found = [l for l in LOCKED + (SWITCH,) if l in raw]
        targets = [_rel(root, cwd, t) for t in found]
    else:
        return None
    if any(t == SWITCH for t in targets):
        return f'组件维护开关 {SWITCH} 只能由用户本人修改，AI 不得写入。'
    if any(t == APPROVALS for t in targets):
        return f'组件维护批准记录 {APPROVALS} 只由 Hook 根据用户本人的消息写入，AI 不得写入。'
    if maintenance_on(root) or approved(root, payload):
        return None
    bad = sorted({t for t in targets if _locked(t, include_switch=False)})
    if not bad:
        return None
    return ('组件只读：搜索组件、投递组件的代码和 Hook 配置已锁定，本会话未获用户批准。'
            '需要修改时说明要改什么、为什么，请用户在对话里明确批准（如“批准修改组件”），'
            f'或由用户本人把 {SWITCH} 的“开启”改为 true。被拦截：' + '、'.join(bad[:5]))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--workspace', required=True)
    ap.add_argument('--host', choices=['claude', 'codex', 'agy'], required=True)
    ap.add_argument('--event')
    a = ap.parse_args()
    root = Path(a.workspace).resolve()
    raw = sys.stdin.read()
    try:
        payload = json.loads(raw) if raw.strip() else {}
    except ValueError:
        payload = {}
    event = a.event or payload.get('hook_event_name') or ''
    note = record_approval(root, payload) if event == 'UserPromptSubmit' else ''
    # 用户只关闭“写入 Excel 提醒”：该会话不再提示和拦截写入，但门禁规则、浏览器规则照常注入。
    reminders = sync_reminders_enabled(root, payload)
    if event in ('Stop', 'SubagentStop') and not reminders:
        emit({})
        return 0
    cfg = load_cfg(root)
    if cfg is None:
        emit({})
        return 0
    if event in ('SessionStart', 'UserPromptSubmit', 'SubagentStart', 'PreInvocation'):
        extra = pending_text(root)[0] if reminders else ''
        ctx = rules_text(cfg, root) + browser_rules(root) + note + extra
        if a.host == 'agy':
            emit({'injectSteps': [{'ephemeralMessage': ctx}]})
        else:
            emit({'hookSpecificOutput': {'hookEventName': event, 'additionalContext': ctx}})
        return 0
    if event == 'PreToolUse':
        hits = tool_violation(cfg, payload, root)
        reason = (excel_violation(payload) or component_violation(root, payload) or handoff_violation(root, payload)
                  or stack_violation(root, payload))
        if hits and not reason:
            reason = ('投递状态门禁：禁止绕过固定脚本直接修改 ' + '、'.join(hits)
                      + '。状态变更请走 json_excel_sync.py propose→用户确认→confirm；写 Excel 只能用 json_excel_sync.py write。')
        if reason:
            if a.host == 'agy':
                emit({'decision': 'deny', 'reason': reason})
            else:
                emit({'hookSpecificOutput': {'hookEventName': 'PreToolUse', 'permissionDecision': 'deny', 'permissionDecisionReason': reason}})
            return 0
        emit({'decision': 'allow'} if a.host == 'agy' else {})
        return 0
    if event in ('Stop', 'SubagentStop'):
        if payload.get('stop_hook_active'):
            emit({})
            return 0
        extra, p = pending_text(root)
        quota = sync(root, 'check')
        notes = []
        if extra:
            notes.append(extra.strip())
        if quota.get('quota_full_but_unclosed'):
            notes.append('名额已被用户选满但仍有岗位未收口：' + '；'.join(quota['quota_full_but_unclosed'][:5]))
        if notes:
            emit({'decision': 'block', 'reason': '结束前必须处理：' + ' '.join(notes)
                  + ' 请直接问用户：“JSON 有未写入 Excel 的改动，现在写入，还是下一轮再写入？”然后结束本轮等待回答。'})
            return 0
        emit({})
        return 0
    emit({})
    return 0


if __name__ == '__main__':
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    try:
        sys.exit(main())
    except Exception as exc:  # Hook 自身故障不阻断宿主，但要可见
        print('投递状态Hook异常：' + str(exc), file=sys.stderr)
        emit({})
        sys.exit(0)
