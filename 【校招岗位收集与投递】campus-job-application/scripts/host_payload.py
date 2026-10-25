"""宿主事件适配：把 Claude / Codex / DSH / Antigravity 各自的 Hook payload 归一成门禁需要的字段。

门禁核心（status_sync_hook.py）只依赖这里导出的函数：取工具名与参数（含 Codex exec 包装里的单个 apply_patch）、
取会话标识、取用户本人发出的原话（含 Codex 问答回复格式）。新增宿主只改本文件，不改门禁规则。
"""
import json
import re


def call_parts(payload):
    call = payload.get('toolCall') or {}
    tool = str(payload.get('tool_name') or payload.get('toolName') or call.get('name') or '')
    ti = payload.get('tool_input')
    if ti is None:
        ti = payload.get('toolInput') or payload.get('arguments') or call.get('arguments') or {}
    # A standalone orchestration wrapper still has an exact patch target. Parse
    # only one literal call; mixed scripts retain normal shell write protection.
    code = ti if isinstance(ti, str) else next((ti[k] for k in ('code', 'input', 'source') if isinstance(ti.get(k), str)), '') if isinstance(ti, dict) else ''
    if re.search(r'(?:^|[._])exec$', tool) and isinstance(code, str):
        wrapped = re.fullmatch(r'\s*(?:(?:text|notify)\s*\(\s*)?await\s+tools\.apply_patch\(("(?:\\.|[^"\\])*")\)\s*\)?\s*;?\s*', code, re.S)
        if wrapped:
            try:
                return 'apply_patch', json.loads(wrapped.group(1))
            except (ValueError, TypeError):
                pass
    return tool, ti


def _session(payload):
    for k in ('session_id', 'sessionId', 'conversation_id', 'conversationId', 'thread_id', 'threadId'):
        if payload.get(k):
            return str(payload[k])
    return '_no_session'


PROMPT_KEYS = ('prompt', 'user_prompt', 'userPrompt', 'message', 'input', 'text', 'content')


def _prompt(payload):
    """各宿主放用户原话的字段名不同：按常见字段名逐层找第一个非空字符串。"""
    stack = [payload]
    while stack:
        cur = stack.pop(0)
        if isinstance(cur, dict):
            for k in PROMPT_KEYS:
                v = cur.get(k)
                if isinstance(v, str) and v.strip():
                    return v
            stack += [v for v in cur.values() if isinstance(v, (dict, list))]
        elif isinstance(cur, list):
            stack += [v for v in cur if isinstance(v, (dict, list, str))]
        elif isinstance(cur, str) and cur.strip():
            return cur
    return ''


QUESTION_REPLY_TAG = 'send_user_message_question_reply'
QUESTION_REPLY = re.compile(r'\A<send_user_message_question_reply>\s*(.*?)\s*</send_user_message_question_reply>\Z', re.S)
QUESTION_CALL_ID = re.compile(r'call_[A-Za-z0-9_-]+\Z')


def _approval_prompts(payload):
    """Read only the submitted user's text; quoted questions cannot grant or deny approval."""
    prompt = next((payload[k] for k in PROMPT_KEYS
                   if isinstance(payload.get(k), str) and payload[k].strip()), '')
    if not prompt:
        return []
    if QUESTION_REPLY_TAG not in prompt:
        return [prompt]
    match = QUESTION_REPLY.fullmatch(prompt.strip())
    if not match:
        return []

    def unique_object(pairs):
        value = {}
        for key, item in pairs:
            if key in value:
                raise ValueError('duplicate key')
            value[key] = item
        return value

    try:
        items = json.loads(match.group(1), object_pairs_hook=unique_object)
        if not isinstance(items, list) or not items:
            return []
        answers, seen = [], set()
        for item in items:
            if (not isinstance(item, dict)
                    or set(item) != {'questionItemId', 'question', 'answer'}
                    or not isinstance(item['questionItemId'], str)
                    or not isinstance(item['question'], str)
                    or not isinstance(item['answer'], str)):
                return []
            identity = json.loads(item['questionItemId'])
            if (not isinstance(identity, list) or len(identity) != 3
                    or identity[0] != 'request_user_input_async'
                    or not isinstance(identity[1], str)
                    or not QUESTION_CALL_ID.fullmatch(identity[1])
                    or type(identity[2]) is not int or identity[2] < 0):
                return []
            key = (identity[1], identity[2])
            if key in seen:
                return []
            seen.add(key)
            answers.append(item['answer'])
        return answers
    except (ValueError, TypeError, RecursionError):
        return []
