#!/usr/bin/env python3
"""投递状态与 JSON→Excel Hook：写入保护全局生效，同步提醒可按对话关闭。

用法：status_sync_hook.py --workspace <工作区根> --host claude|codex|agy [--event <宿主事件名>]
stdin 为宿主 Hook payload。

- 会话/每轮开始：注入标签体系与门禁规则；若 JSON 有未写入 Excel 的改动且用户未选择延后，要求先询问用户。
- 工具调用前：拦截绕过固定脚本直接修改 Excel、主数据 JSON、确认台账、标签配置的写操作。
- 对话结束（Stop）：存在未写入且用户未选择延后的改动时阻止结束，要求询问“现在写入还是下一轮写入”。
"""
import argparse
import contextlib
import hashlib
import tempfile
import time
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
from hook_file_policy import (file_effects, path_at, touches, mentioned, text_only, PATH_KEYS,
                              blank_data_redirects, expand_vars, note_assignments)

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
    base = (cfg.get('companyStore') or {}).get('dir', '岗位交接/公司')
    profile = os.environ.get('RECRUITING_PROFILE_JSON') or cfg.get('profileFile', '产品管理/网申个人信息主档.json')
    # 公司宽表文件只由 company_store.py 生成；投递回传只由投递组件写入、company_store.py 消费
    return [cfg['workbook'], cfg['masterData'], cfg['confirmationLedger'], LABELS_FILE, *cfg.get('labelGroups', []),
            base + '/', str(Path(base).with_name('投递回传.json')), '产品管理/user_applicaiton_requirements_temp.json',
            '.runtime/user_application_requirements_state.json', profile]


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
            + f'2. 数据文件保护统一遵循 {Path(__file__).resolve().parent.parent / "references/data-file-protection.md"}，执行前读取该唯一规则来源。\n'
            '3. 任何状态变更：先 propose，把输出的“当前状态是X，更新状态是Y”原样给用户看，取得用户明确确认后才能 confirm（附用户原话）。行一律按编码引用（岗位编码 J000001；公司行=地区表+公司编码 C00001）。岗位链接、投递链接、站点的登录链接/账号核对页只能用 relink / set-apply-link / set-site-links（先 --preview 给用户看，确认后带原话执行），不能用 upsert 直接改。官网无独立详情页的岗位，链接例外备注只能用 link-exception 写（先 --preview 给用户看，确认后执行）。\n'
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
    _, docs = st.round()
    handed = {c: d for c, d in docs.items() if d['批次']['选中岗位']}
    if not handed:
        return '正向门禁：投递批次未经用户正向评审确认。请先运行 company_store.py forward 生成评审页，用户确认后 forward-confirm。'
    for c, d in handed.items():
        problem = company_store.verify_handoff(d)
        if problem:
            return f"正向门禁：{d['公司名称']}（{c}）{problem}。请先运行 company_store.py forward 生成评审页，用户确认后 forward-confirm。"
    return None


PATCH_TARGET = re.compile(r'^\*\*\* (?:(?:Update|Add|Delete) File|Move to): *(.+?)\s*$', re.M)


def write_targets(ti):
    """编辑类调用真正改写的文件：apply_patch 取补丁头里的目标路径，Write/Edit 取路径参数；无法确定时返回 None。"""
    texts = [ti] if isinstance(ti, str) else [v for v in ti.values() if isinstance(v, str)] if isinstance(ti, dict) else []
    for t in texts:
        if '*** Begin Patch' in t:
            return PATCH_TARGET.findall(t)
    if isinstance(ti, dict):
        found = [ti[k] for k in PATH_KEYS if isinstance(ti.get(k), str) and ti[k]]
        if found:
            return found
    return None


def data_effects(cfg, payload, root):
    """Consume the framework's existing protected(cfg) list; no shadow manifest."""
    registered = [str(path_at(p.rstrip('/'), root)) + ('/' if p.endswith('/') else '') for p in protected(cfg)]
    script = path_at(cfg.get('syncScript') or str(SCRIPT), root)
    try:
        return file_effects(payload, root, registered, script)
    except (ValueError, TypeError, OSError, SyntaxError) as error:
        _, args = call_parts(payload)
        return mentioned(args, payload.get('cwd') or root, registered, '工具解析失败：' + type(error).__name__)


def data_reason(rule, detail, action):
    return f'【数据文件保护：{rule}】Hook：{Path(__file__).resolve()}。{detail}。正确做法：{action}'


def excel_violation(payload, cfg=None, root=None):
    """The registered workbook is an object, not a file-extension pattern."""
    if root is None or not cfg or not cfg.get('workbook'):
        return None
    book = path_at(cfg['workbook'], root)
    hit = next((e for e in data_effects(cfg, payload, root) if touches(e, book)), None)
    if hit:
        return data_reason('登记工作簿', str(book) + '：' + (hit.reason or '不允许直接访问工作簿内容'),
                           '读取主数据 JSON；同步工作簿使用配置登记的同步脚本。数据保护不处理 Git 命令')
    return None


# 第 2 步完成前，原公司文件夹里的 登录/表单定义 配置仍在原位（只读）；公司宽表一家一份，以公司编码命名。
HANDOFF_ALLOWED = re.compile(r'^岗位交接/(工作簿全量\.json|正向评审\.html|投递回传\.json|反向评审\.html|公司/C\d{5}\.json|公司/[^/]+/[^/]+/(登录(@[^/]+)?|表单定义(@[^/]+)?)\.json)$')
STACK_NAME = re.compile(r'(^|/)(历史|备份|快照|存档)(/|$)|上一版|副本|[-_.]v\d+\.[A-Za-z]+$|\.(bak|orig)$')
PATH_TOKEN = re.compile(r'[^\s\'"<>|;&()]*(?:岗位交接/|历史/|备份|快照|上一版|副本|\.bak)[^\s\'"<>|;&()]*')


def stack_violation(root, payload):
    """记录唯一门禁：同一信息只保留最新一版、覆盖更新；不得在工作区另存副本/历史/备份，岗位交接/ 只允许固定文件。"""
    root = Path(root).resolve()
    # Record uniqueness is a separate creation rule, not a shadow data list.
    targets = [str(e.path) for e in file_effects(payload, root) if e.kind == 'write']
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
    return data_reason('记录唯一门禁', ('同一信息只保留最新一版、覆盖原文件更新，禁止另存副本/历史/备份/按批次另起同类文件（回滚靠 git）。'
            '岗位交接/ 下只允许 工作簿全量.json、正向评审.html、投递回传.json、反向评审.html、公司/<公司编码>.json（以及迁出前的 公司/<地区>/<公司>/{登录[@域名].json,表单定义[@域名].json}）；'
            '被拦截：' + '、'.join(bad[:5])), '更新唯一原文件；过程文件放工作区外临时目录，用完删除')


def tool_violation(cfg, payload, root=None):
    tool, ti = call_parts(payload)
    raw = json.dumps(ti, ensure_ascii=False)
    files = protected(cfg)
    if root is not None:
        effects = [e for e in data_effects(cfg, payload, root) if e.kind in {'write', 'delete', 'unknown'}]
        return [f for f in files if any(touches(e, path_at(f.rstrip('/'), root))
                or f.endswith('/') and path_at(f.rstrip('/'), root) in e.path.parents for e in effects)] or None
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
                    target, guarded = p.resolve(), (Path(root).resolve() / f).resolve()
                    return target == guarded or f.endswith('/') and guarded in target.parents
                except OSError:
                    pass
            t = t.replace('\\', '/')
            return t == f or t.endswith('/' + f) or f.endswith('/') and (t.startswith(f) or '/' + f in t)
        return [f for f in files if any(hit(t, f) for t in targets)] or None
    if SHELL_TOOL.search(tool) and not EDIT_TOOL.search(tool) or isinstance(ti, dict) and 'command' in ti:
        hits = [f for f in files if f in raw or Path(f).name in raw]
        # 与组件软锁同一套写入判断：重定向、cp/mv/sed -i 等改写命令、python/node 内联写文件都算写入。
        cmd = ti.get('command') if isinstance(ti, dict) and isinstance(ti.get('command'), str) else raw
        writes = bool(SHELL_WRITE.search(raw) or INLINE_WRITE.search(raw))
        if root is not None and not writes:
            rels = _shell_targets(Path(root).resolve(), payload.get('cwd') or str(root), cmd)
            writes = any(t and any(t == f.rstrip('/') or (f.endswith('/') and t.startswith(f)) for f in files) for t in rels)
        if hits and 'json_excel_sync.py' not in raw and writes:
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
            hits = [f for f in files if any(t == f or t.endswith('/' + f) or (f.endswith('/') and f in t) or (not f.endswith('/') and Path(t).name == Path(f).name) for t in targets)]
        return hits or None
    return None


# 组件软锁：两个组件的代码与 Hook 配置默认锁定。两种放开方式：
#   1) 用户在对话里明确批准修改组件（如“批准修改组件”“授权你改 Hook”），本会话放开，用户说“收回授权/锁上组件”或 12 小时后失效；
#      批准只从用户本人发出的消息（UserPromptSubmit）识别，由 Hook 记录到 APPROVALS，AI 不能写这个记录。
#   2) 用户本人把维护开关文件的“开启”改为 true（长期放开）。开关文件与批准记录 AI 任何时候都不能写。
LOCKED = ('.agents/skills/campus-job-application/', '.agents/skills/job-application-assistant-v2/',
          '.agents/skills/campus-job-plan-inventory/', '.agents/skills/campus-job-preference-intake/',
          '.claude/settings.json', '.codex/hooks.json')
SWITCH = '产品管理/组件维护.json'
APPROVALS = '.runtime/组件维护批准.json'
APPROVAL_HOURS = 12
COMPONENT_WORD = re.compile(r'组件|插件|[Hh]ook|钩子|[Ss]kill')
APPROVE_MODIFY = re.compile(r'(批准|授权|同意|允许).{0,6}(修改|改|维护|更新)')
CONDITIONAL = re.compile(r'情况下|的话|前提|时才|之后才|需要.{0,4}(批准|授权)|是否')
NEGATED = re.compile(r'(不|别|禁止|不要|不准|不许)\s*(批准|授权|同意|允许|修改|改|维护)')
REVOKE = re.compile(r'(收回|撤销|撤回|取消|结束)\s*(组件)?\s*(维护)?\s*(授权|批准|维护)|锁上组件|锁回组件|恢复只读')
SEGMENT = re.compile(r'&&|\|\||;|\n|(?<!>)\|')  # `>|` 是重定向，不是管道
REDIRECT = re.compile(r'(?<!>)>{1,2}(\|?&?)\s*([^\s&;|>][^\s;&|]*)')  # 含 1> 2> &> >| >&
DESCRIPTOR = re.compile(r'\d+-?|-')  # `>&` 后面是这些时只是复制或关闭描述符，不写文件
REDIRECT_TOKEN = re.compile(r'(?:\d+|&)?>')
BARE_REDIRECT = re.compile(r'(?:\d*|&)>{1,2}\|?&?')  # 整个词就是操作符，目标在下一个词
LN_TARGET_DIR = re.compile(r'(?:--t[a-z-]*=?|-[bdFfiLnPrsTv]*t)(.*)')
IN_PLACE = {'sed', 'perl', 'truncate', 'touch', 'chmod', 'chown', 'rm', 'rmdir', 'unlink', 'patch', 'install', 'tee', 'mkdir'}
COPY = {'cp', 'mv', 'rsync', 'ditto'}
GIT_MUTATE = {'checkout', 'restore', 'reset', 'apply', 'am', 'stash', 'clean', 'pull', 'merge', 'rebase', 'cherry-pick', 'rm', 'mv', 'switch', 'revert'}
PKG_MUTATE = {'install', 'i', 'ci', 'uninstall', 'remove', 'update', 'add', 'upgrade'}
INLINE_WRITE = re.compile(r'(write_text|write_bytes|open\([^)]*["\'][wax]b?\+?["\']|json\.dump\(|writeFile|appendFile|unlink|rmSync|shutil\.|os\.remove|os\.rename|os\.replace|renameSync|copyFile|cpSync|createWriteStream|\.rename\(|truncateSync|openSync\([^)]*["\'][wa]\+?["\'])')


def maintenance_on(root):
    try:
        return json.loads((root / SWITCH).read_text(encoding='utf-8')).get('开启') is True
    except (OSError, ValueError, AttributeError):
        return False


def _debug(root, event, payload, found, **details):
    """Keep the latest approval write and permission check in one diagnostic."""
    target = root / '.runtime/组件软锁诊断.json'
    record = {
        'event': event, 'keys': sorted(payload)[:40], 'session': _session(payload),
        'workspace': str(Path(root).resolve()), 'source': str(Path(__file__).resolve()),
        'approvalPath': str((root / APPROVALS).resolve()), 'promptFound': found,
        'at': time.strftime('%Y-%m-%d %H:%M:%S'), **details,
    }
    try:
        with _approval_lock(root, name='.组件软锁诊断.lock'):
            try:
                previous = json.loads(target.read_text(encoding='utf-8'))
                if not isinstance(previous, dict):
                    previous = {}
            except (OSError, ValueError):
                previous = {}
            latest_approval = record if 'saved' in details else previous.get('lastApproval')
            latest_check = record if details.get('operation') == 'check' else previous.get('lastCheck')
            body = {**record, 'lastApproval': latest_approval, 'lastCheck': latest_check}
            _replace_json(target, body)
    except OSError:
        pass


def _load_approvals(root):
    try:
        d = json.loads((root / APPROVALS).read_text(encoding='utf-8'))
        return d if isinstance(d, dict) else {}
    except (OSError, ValueError):
        return {}


# An independent question after a real approval is not a condition on that
# approval. Keep commas/newlines in the same condition scope so splitting
# "如果...，批准..." cannot turn a conditional instruction into permission.
APPROVAL_CONDITION = re.compile(r'如果|若(?!干)|假如|倘若|仅当|只有|除非|前提|条件|满足.{0,20}后|待.{0,20}(?:后|再)|等.{0,20}(?:后|再)')
# A rule describes who may approve ("用户...批准...后"); a person approving says it directly.
RULE_SUBJECT = re.compile(r'^(?:由|经|须|需|请|待|等|当|在|若)?\s*用户')
AFTER_THE_FACT = re.compile(r'后\s*[。；;！!]?$')
APPROVAL_POST_CONDITION = re.compile(r'^\s*(?:(?:但|不过|且|同时|仅|只)\s*)?(?:前提|条件|必须|须|仅当|只有|除非|如果|若|待|等.{0,20}(?:后|再))')


def direct_authority_text(prompt):
    """Quoted examples and forwarded blocks are data, never an instruction."""
    text = re.sub(r'(?ms)^\s*(`{3,}|~{3,})[^\n]*\n.*?^\s*\1[^\n]*', '\n', prompt)
    text = re.sub(r'(?ms)^\s*(?:`{3,}|~{3,}).*$', '', text)
    kept, quoted = [], False
    for line in text.splitlines():
        if not line.strip():
            quoted = False; kept.append(line); continue
        if re.match(r'^\s*>', line):
            quoted = True
        if not quoted: kept.append(line)
    text = '\n'.join(kept)
    text = re.sub(r'`[^`]*`', '', text, flags=re.S)
    for opening, closing in [('“', '”'), ('‘', '’'), ('「', '」'), ('『', '』'), ('"', '"'), ("'", "'")]:
        text = re.sub(re.escape(opening) + r'.*?' + re.escape(closing), '', text, flags=re.S)
    return text


def reported_authority(clause, action):
    prefix = clause[:action.start()]
    return bool(re.search(r'(?:说|回复|输入|发送|写道|写下|告诉|提到|引用|转述|转贴|转发|原话|示例|例如|比如).{0,24}$', prefix))


def direct_revocation(prompt):
    text = re.sub(r'\r?\n+', '，', direct_authority_text(prompt))
    for sentence in re.split(r'(?<=[。；;！？!?])', text):
        if sentence.rstrip().endswith(('?', '？')):
            continue
        clauses = [part.strip() for part in re.split(r'[,，]', sentence) if part.strip()]
        for index, clause in enumerate(clauses):
            match = REVOKE.search(clause)
            if not match or reported_authority(clause, match):
                continue
            prefix = '，'.join(clauses[:index]) + clause[:match.start()]
            if re.search(r'不|别|禁止|不要|是否', prefix) or APPROVAL_CONDITION.search(prefix + clause) or re.search(r'(?:吗|么)\s*$', clause):
                continue
            return True
    return False


def component_approval_quotes(prompt):
    """Return only unconditional, non-question authorization clauses.

    This consumes the existing host adapter's submitted user/answer text; it
    does not parse UI questions, change session identity or register approval.
    """
    prompt = direct_authority_text(prompt)
    quotes = []
    # Newlines do not break a leading condition. Strong sentence punctuation
    # still lets a later unrelated question coexist with an earlier approval.
    sentences = re.split(r'(?<=[。；;！？!?])', re.sub(r'\r?\n+', '，', prompt))
    carry_condition = False
    for sentence in sentences:
        sentence = sentence.strip()
        if not sentence:
            continue
        clauses = [part.strip() for part in re.split(r'[,，]', sentence) if part.strip()]
        # A direct component denial must not be silently overridden by a
        # positive clause in the same submitted instruction.
        if any(COMPONENT_WORD.search(c) and NEGATED.search(c) for c in clauses):
            quotes = []
            carry_condition = False
            continue
        matched = False
        for index, clause in enumerate(clauses):
            if not (COMPONENT_WORD.search(clause) and APPROVE_MODIFY.search(clause)):
                continue
            action = APPROVE_MODIFY.search(clause)
            if reported_authority(clause, action) or RULE_SUBJECT.search(clause) or AFTER_THE_FACT.search(clause):
                continue
            matched = True
            prefix = '，'.join(clauses[:index])
            suffix = clauses[index + 1:]
            question = clause.rstrip().endswith(('?', '？')) or re.search(r'(?:吗|么)\s*[。；;！!]?$', clause)
            scoped_condition = (carry_condition or APPROVAL_CONDITION.search(clause)
                                or APPROVAL_CONDITION.search(prefix)
                                or any(APPROVAL_POST_CONDITION.search(c) for c in suffix))
            if scoped_condition:
                # A later conditional permission narrows an earlier broad
                # approval; do not silently keep the earlier unrestricted one.
                quotes = []
                continue
            if (question or NEGATED.search(clause) or CONDITIONAL.search(clause)
                    or CONDITIONAL.search(prefix)):
                continue
            quotes.append(clause.rstrip('。；;！!').strip())
        # Conservatively retain a standalone condition fragment across a full
        # stop. A completed conditional instruction/question does not poison an
        # unrelated later explicit approval.
        carry_condition = bool(not matched and len(clauses) == 1
                               and APPROVAL_CONDITION.match(sentence)
                               and not sentence.endswith(('?', '？')))
    return quotes


def without_injected_rules(root, prompt):
    """DSH 把排队的 SessionStart 注入和用户消息拼进同一个 prompt：逐行去掉本 Hook 自己注入的规则。"""
    cfg = load_cfg(root)
    if cfg is None:
        return prompt
    try:
        injected = {line.strip() for line in (rules_text(cfg, root) + browser_rules(root)).splitlines() if line.strip()}
    except Exception:  # 规则生成失败时仍有 RULE_SUBJECT / AFTER_THE_FACT 兜底
        return prompt
    return '\n'.join(line for line in prompt.splitlines() if line.strip() not in injected)


@contextlib.contextmanager
def _approval_lock(root, name='.组件维护批准.lock'):
    """Serialize the complete read/change/replace across every host process."""
    directory = root / '.runtime'
    directory.mkdir(parents=True, exist_ok=True)
    with (directory / name).open('a+b') as file:
        try:
            import fcntl
            operation = lambda unlock=False: fcntl.flock(file, fcntl.LOCK_UN if unlock else fcntl.LOCK_EX | fcntl.LOCK_NB)
        except ImportError:
            import msvcrt
            if file.seek(0, 2) == 0:
                file.write(b'\0'); file.flush()
            def operation(unlock=False):
                file.seek(0)
                msvcrt.locking(file.fileno(), msvcrt.LK_UNLCK if unlock else msvcrt.LK_NBLCK, 1)
        deadline = time.monotonic() + 2
        while True:
            try:
                operation()
                break
            except OSError as error:
                if error.errno not in (11, 13, 35) or time.monotonic() >= deadline:
                    raise
                time.sleep(.02)
        try:
            yield
        finally:
            operation(True)


def _replace_json(target, data):
    body = json.dumps(data, ensure_ascii=False, indent=2).encode('utf-8')
    descriptor, temporary = tempfile.mkstemp(prefix=f'.{target.stem}.', suffix='.tmp', dir=target.parent)
    try:
        with os.fdopen(descriptor, 'wb') as stream:
            stream.write(body)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, target)
        return body
    finally:
        Path(temporary).unlink(missing_ok=True)


def _save_approvals(root, data):
    """Publish one complete file, then verify those bytes before reporting success."""
    target = root / APPROVALS
    body = _replace_json(target, data)
    saved = target.read_bytes()
    if saved != body:
        raise OSError('批准记录写入后读回不一致')
    return hashlib.sha256(saved).hexdigest()


def record_approval(root, payload):
    """Only the host's user event may persist approval; ordinary reads never renew it."""
    prompts = _approval_prompts(payload)
    _debug(root, 'UserPromptSubmit', payload, any(p.strip() for p in prompts))
    prompts = [without_injected_rules(root, p) for p in prompts]
    if not any(p.strip() for p in prompts):
        return ''
    sid = _session(payload)
    approval_quotes = [quote for prompt in prompts for quote in component_approval_quotes(prompt)]
    revoked = any(direct_revocation(p) for p in prompts)
    if not revoked and not approval_quotes:
        return '\n【组件软锁】本会话已获用户批准修改组件。' if approved(root, payload) else ''
    try:
        with _approval_lock(root):
            now = time.time()
            data = {k: v for k, v in _load_approvals(root).items() if v.get('expiresAt', 0) > now}
            if revoked:
                data.pop(sid, None)
                note = '\n【组件软锁】用户已收回本会话的组件修改批准，组件恢复只读。'
            else:
                hours = 2 if sid == '_no_session' else APPROVAL_HOURS
                data[sid] = {'userQuote': '\n'.join(approval_quotes)[:300], 'approvedAt': now, 'expiresAt': now + hours * 3600}
                note = f'\n【组件软锁】已记录用户批准：本会话可以修改搜索组件、投递组件和 Hook 配置（{hours} 小时内有效，用户说“收回授权”即失效）。修改仍须按用户的具体要求进行，并逐项汇报。'
            digest = _save_approvals(root, data)
            _debug(root, 'UserPromptSubmit', payload, True, operation='revoke' if revoked else 'approve',
                   saved=True, approvalSha256=digest, sessionRecorded=sid in data)
        return note + f'\n【批准保存核对】会话 {sid}；文件 {root / APPROVALS}；读回 SHA256 {digest}。'
    except OSError as error:
        _debug(root, 'UserPromptSubmit', payload, True, saved=False, error=str(error))
        return f'\n【组件软锁】批准记录保存失败，未确认生效：{error}。目标文件：{root / APPROVALS}。'


def approved(root, payload):
    v = _load_approvals(root).get(_session(payload))
    allowed = bool(v) and v.get('expiresAt', 0) > time.time()
    _debug(root, payload.get('hook_event_name') or 'PreToolUse', payload, False,
           operation='check', authorizationSource='conversation_record', sessionRecorded=bool(v), approved=bool(allowed))
    return allowed


def _rel(root, cwd, t):
    t = str(t).strip().strip('\'"')
    if not t or t.startswith('-'):
        return None
    p = Path(t).expanduser()
    p = p if p.is_absolute() else Path(cwd) / p
    try:
        return p.resolve().relative_to(root.resolve()).as_posix()
    except (OSError, RuntimeError, ValueError):
        return None


def _locked(rel, include_switch=True):
    if rel is None:
        return False
    if include_switch and rel == SWITCH:
        return True
    return any(rel == l.rstrip('/') or (l.endswith('/') and rel.startswith(l)) for l in LOCKED)


def _without_redirects(toks):
    """去掉重定向词，免得 `ln a b 2>/dev/null` 的“最后一个参数”变成重定向；重定向目标由 REDIRECT 单独判断。"""
    kept, skip_next = [], False
    for tok in toks:
        if skip_next:
            skip_next = False
            continue
        found = REDIRECT_TOKEN.search(tok)
        if not found:
            kept.append(tok)
            continue
        if found.start():
            kept.append(tok[:found.start()])
        skip_next = bool(BARE_REDIRECT.fullmatch(tok[found.start():]))
    return kept


def _link_location(args):
    """`ln` 只写链接位置：-t 指定的目录，否则是最后一个参数；只给来源时落在当前目录。"""
    sources = []
    for i, arg in enumerate(args):
        target_dir = LN_TARGET_DIR.fullmatch(arg)
        if target_dir:
            return target_dir.group(1) or (args[i + 1] if i + 1 < len(args) else None)
        if not arg.startswith('-'):
            sources.append(arg)
    if len(sources) > 1:
        return sources[-1]
    return Path(sources[0]).name if sources else None


def _shell_targets(root, cwd, cmd):
    """粗略解析 shell 命令里会被改写的路径（相对路径按 cd 后的目录解析）。"""
    out, assigned = [], {}
    if '*** Begin Patch' in cmd:
        out += [_rel(root, cwd, t) for t in PATCH_TARGET.findall(cmd)]
    cmd = blank_data_redirects(cmd, anchors=LOCKED + (SWITCH, APPROVALS))
    for seg in SEGMENT.split(cmd):
        out += [_rel(root, cwd, expand_vars(t.strip('\'"'), assigned)) for op, t in REDIRECT.findall(seg)
                if t != '/dev/null' and not ('&' in op and DESCRIPTOR.fullmatch(t))]
        try:
            toks = shlex.split(seg)
        except ValueError:
            toks = seg.split()
        toks = [expand_vars(t, assigned) for t in toks]
        note_assignments(toks, assigned)
        toks = _without_redirects(toks)
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
        elif name == 'ln':
            location = _link_location(args)
            if location:
                out.append(_rel(root, cwd, location))
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
    name = tool.rsplit('__', 1)[-1].rsplit('.', 1)[-1].lower()
    if text_only(name, ti):
        return None
    if isinstance(ti, dict) and isinstance(ti.get('command'), str) and ti['command'].strip() or SHELL_TOOL.search(tool) and not EDIT_TOOL.search(tool):
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
    if maintenance_on(root):
        _debug(root, payload.get('hook_event_name') or 'PreToolUse', payload, False,
               operation='check', authorizationSource='user_maintenance_switch', approved=True)
        return None
    if approved(root, payload):
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
        data = excel_violation(payload, cfg, root)
        component = component_violation(root, payload)
        handoff = handoff_violation(root, payload)
        unique = stack_violation(root, payload)
        if hits and not data:
            unknown = sorted({e.reason for e in data_effects(cfg, payload, root) if e.kind == 'unknown' and e.reason})
            detail = '；'.join(unknown) if unknown else '禁止绕过固定入口直接改写'
            data = data_reason('登记数据文件', detail + '；登记对象：' + '、'.join(hits),
                               '使用框架固定写入命令；状态变更走 propose→用户确认→confirm；入口配置走 site-entry 预览→用户确认→写入')
        reason = '\n'.join(r for r in (data, component, handoff, unique) if r)
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
            notes.append('名额已被用户选满但仍有岗位未收口：' + '；'.join(quota['quota_full_but_unclosed'][:5])
                         + '。请用 json_excel_sync.py propose 把收口变更原样给用户看并等待确认。')
        if notes:
            ask = ' 请直接问用户：“JSON 有未写入 Excel 的改动，现在写入，还是下一轮再写入？”然后结束本轮等待回答。' if extra else ' 然后结束本轮等待回答。'
            emit({'decision': 'block', 'reason': '结束前必须处理：' + ' '.join(notes) + ask})
            return 0
        emit({})
        return 0
    emit({})
    return 0


if __name__ == '__main__':
    for stream in (sys.stdin, sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            stream.reconfigure(encoding='utf-8')
    try:
        sys.exit(main())
    except Exception as exc:  # Hook 自身故障不阻断宿主，但要可见
        print('投递状态Hook异常：' + str(exc), file=sys.stderr)
        emit({})
        sys.exit(0)
