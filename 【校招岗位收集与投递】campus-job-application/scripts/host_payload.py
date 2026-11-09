"""宿主事件适配：把 Claude / Codex / DSH / Antigravity 各自的 Hook payload 归一成门禁需要的字段。

门禁核心（status_sync_hook.py）只依赖这里导出的函数：取工具名与参数（含 Codex exec 包装里的单个 apply_patch）、
取会话标识、取用户本人发出的原话（含 Codex 问答回复格式）。新增宿主只改本文件，不改门禁规则。
"""
import json
import os
import re
from datetime import datetime
from pathlib import Path


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


def _submitted_prompt(payload):
    if payload.get('hook_event_name', 'UserPromptSubmit') != 'UserPromptSubmit':
        return ''
    if payload.get('role', 'user') != 'user':
        return ''
    source = payload.get('source')
    if isinstance(source, dict):
        source = source.get('kind', source.get('type'))
    if source not in (None, 'user'):
        return ''
    return next((payload[k] for k in PROMPT_KEYS
                 if isinstance(payload.get(k), str) and payload[k].strip()), '')


def _card_answers(prompt):
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
            answers.append((identity[1], identity[2], item['question'], item['answer']))
        return answers
    except (ValueError, TypeError, RecursionError):
        return []


def _approval_prompts(payload):
    """Parse submitted text. Card authorization additionally requires a verified request."""
    prompt = _submitted_prompt(payload)
    if not prompt:
        return []
    if QUESTION_REPLY_TAG in prompt:
        return [answer for _, _, _, answer in _card_answers(prompt)]
    return [prompt]


def codex_sessions_root():
    return (Path(os.environ.get('CODEX_HOME') or Path.home() / '.codex') / 'sessions').resolve()


def codex_card_receipts(root, payload, now, lifetime):
    """Bind each card item to a recent native question in this workspace and session."""
    answers = _card_answers(_submitted_prompt(payload))
    transcript = payload.get('transcript_path')
    if not answers or not isinstance(transcript, str) or not transcript:
        return {}
    try:
        transcript = Path(transcript).resolve(strict=True)
        if not transcript.is_relative_to(codex_sessions_root()) or transcript.suffix != '.jsonl':
            return {}
        with transcript.open(encoding='utf-8') as stream:
            meta = json.loads(stream.readline())
            session = meta['payload']
            if (meta['type'] != 'session_meta' or session['id'] != _session(payload)
                    or not Path(session['cwd']).resolve().is_relative_to(Path(root).resolve())):
                return {}
            requests = {}
            wanted = {call for call, _, _, _ in answers}
            for line in stream:
                if '"function_call"' not in line or not any(call in line for call in wanted):
                    continue
                item = json.loads(line)
                call = item.get('payload', {})
                call_id = call.get('call_id')
                if (item.get('type') != 'response_item' or call.get('type') != 'function_call'
                        or call.get('name') != 'request_user_input_async' or call_id not in wanted):
                    continue
                issued = datetime.fromisoformat(item['timestamp'].replace('Z', '+00:00')).timestamp()
                if not 0 <= now - issued < lifetime or call_id in requests:
                    return {}
                requests[call_id] = (json.loads(call['arguments'])['questions'], issued + lifetime)
        receipts = {}
        for call_id, index, question, _ in answers:
            questions, expires = requests[call_id]
            if questions[index]['title'] != question:
                return {}
            receipts[f'{call_id}:{index}'] = expires
        return receipts
    except (OSError, ValueError, TypeError, KeyError, IndexError, OverflowError, AttributeError):
        return {}
