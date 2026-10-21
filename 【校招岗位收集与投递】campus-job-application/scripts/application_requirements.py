#!/usr/bin/env python3
"""一次性投递需求：覆盖写入、成功组批后消费，同回合默认复用、新回合初始化。"""
import argparse
from contextlib import contextmanager
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import sys
import time
import uuid

TEMP = '产品管理/user_applicaiton_requirements_temp.json'
STATE = '.runtime/user_application_requirements_state.json'


def save(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name('.' + path.name + '.' + uuid.uuid4().hex + '.tmp')
    try:
        tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
        os.replace(tmp, path)
    finally:
        if tmp.exists():
            tmp.unlink()


def read(path, default=None):
    if not path.exists():
        return {} if default is None else default
    try:
        value = json.loads(path.read_text(encoding='utf-8'))
    except (ValueError, OSError) as exc:
        raise ValueError('需求记录无法读取，停止以免误走默认路线') from exc
    if not isinstance(value, dict):
        raise ValueError('需求记录必须是JSON对象')
    return value


def normalize(requirements):
    if not isinstance(requirements, dict):
        raise ValueError('用户需求必须是JSON对象')
    if set(requirements) - {'公司', '地区', '浏览器'}:
        raise ValueError('不支持的需求参数；当前支持 公司、地区、浏览器')
    out = dict(requirements)
    companies = out.get('公司')
    if companies is None:
        out.pop('公司', None)
    if isinstance(companies, str):
        companies = [companies]
    if companies is not None:
        if not isinstance(companies, list) or any(not isinstance(x, str) or not x.strip() for x in companies):
            raise ValueError('用户需求.公司须为公司名称列表')
        companies = [x.strip() for x in companies]
        if len(companies) != len(set(companies)):
            raise ValueError('用户需求.公司存在重复名称')
        if companies:
            out['公司'] = companies
        else:
            out.pop('公司', None)
    if out.get('地区') is not None and (not isinstance(out['地区'], str) or not out['地区'].strip()):
        raise ValueError('地区必须是非空字符串')
    if '浏览器' in out and out['浏览器'] not in ('iab', 'chrome', 'dual'):
        raise ValueError('浏览器只能是 iab、chrome、dual')
    return out


def timestamp(value):
    if isinstance(value, dict):
        value = value.get('value')
    if not value:
        return float('-inf')
    try:
        dt = datetime.fromisoformat(str(value).replace('Z', '+00:00'))
        return dt.replace(tzinfo=timezone.utc).timestamp() if dt.tzinfo is None else dt.timestamp()
    except (ValueError, TypeError, OverflowError):
        return float('-inf')


def select_companies(rows, ledger, requirements):
    """公司以地区+名称为身份；保留一家公司所有待投递岗位，不按岗位数截断。"""
    import json_excel_sync as S
    req = normalize(requirements)
    confirmations = {S.key_id(e['key']): timestamp(e.get('confirmed_at'))
                     for e in ledger.get('entries', [])
                     if e.get('key') and e.get('to') == '岗位待投递'}
    pending = {}
    for key, records in rows.items():
        if req.get('地区') and key[0] != req['地区']:
            continue
        roles = [r for r in records if r['kind'] == 'role' and r['values'].get('投递状态') == '岗位待投递']
        if roles:
            pending[key] = roles
    if req.get('公司'):
        selected = []
        for name in req['公司']:
            matches = [key for key in pending if key[1] == name]
            if not matches:
                raise ValueError('指定公司没有待投递岗位：' + name)
            if len(matches) != 1:
                raise ValueError('同名公司跨地区，须指定地区：' + name)
            selected.append(matches[0])
        return selected
    def latest(key):
        dates = []
        for r in pending[key]:
            date = confirmations.get(S.key_id(r['key']), float('-inf'))
            dates.append(timestamp(r['values'].get('数据刷新时间')) if date == float('-inf') else date)
        return max(dates)
    # stable sort: identical dates and undated companies retain their master-table order.
    return sorted(pending, key=latest, reverse=True)[:10]


class Requirements:
    def __init__(self, root, session=None, turn=None):
        self.root = Path(root).resolve()
        self.temp = self.root / TEMP
        self.state = self.root / STATE
        self.session = session or os.environ.get('DSH_SESSION_ID') or os.environ.get('CODEX_THREAD_ID') or os.environ.get('CLAUDE_SESSION_ID')
        self.turn = turn or os.environ.get('CODEX_TURN_ID') or os.environ.get('DSH_TURN_ID')

    @contextmanager
    def locked(self):
        lock = self.root / '.runtime/user_application_requirements.lock'
        lock.parent.mkdir(parents=True, exist_ok=True)
        deadline = time.monotonic() + 2
        while True:
            try:
                fd = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
                os.close(fd)
                break
            except FileExistsError:
                if time.monotonic() >= deadline:
                    raise ValueError('需求表正在被另一操作使用，请稍后重试')
                time.sleep(0.02)
        try:
            yield
        finally:
            lock.unlink()

    def context(self, state):
        if not self.session:
            raise ValueError('缺少宿主会话身份')
        current = state.setdefault('sessions', {}).get(self.session, {})
        turn = self.turn
        if not turn:
            # The existing workflow hook supplies genuine host turn metadata for first use.
            f = self.root / '.codex/workflow-state/injections.jsonl'
            if f.exists():
                for line in reversed(f.read_text(encoding='utf-8').splitlines()):
                    event = json.loads(line)
                    if event.get('session_id') == self.session and event.get('turn_id'):
                        if event.get('at', 0) >= current.get('observedAt', 0):
                            turn = event['turn_id']
                        break
            turn = turn or current.get('turnId')
        if not turn:
            raise ValueError('缺少当前回合身份，请先加载需求入口Hook')
        if current.get('turnId') != str(turn):
            current = {'turnId': str(turn), 'phase': 'initial', 'active': current.get('active', False),
                       'observedAt': time.time()}
            state['sessions'][self.session] = current
        return current

    def document(self, current):
        doc = read(self.temp)
        if not doc:
            return {}
        if doc.get('sessionId') != self.session:
            raise ValueError('需求表由另一个对话持有，禁止覆盖或消费')
        if doc.get('turnId') != current['turnId']:
            return {}  # A previous turn cannot become a new turn's implicit parameters.
        normalize(doc.get('requirements', {}))
        return doc

    def begin(self):
        with self.locked():
            state = read(self.state)
            current = self.context(state)
            doc = self.document(current)
            ask = current.get('phase') == 'initial' and not doc
            current['active'] = True
            if ask:
                current['phase'] = 'queried'
            if not self.temp.exists():
                save(self.temp, {})
            save(self.state, state)
            return {'askUserRequirements': ask, 'requirements': doc.get('requirements', {}),
                    'route': 'parameters' if doc else 'default_latest_10', 'turnId': current['turnId']}

    def set(self, requirements, quote):
        req = normalize(requirements)
        if not quote or not quote.strip():
            raise ValueError('需求写入须附用户原话')
        with self.locked():
            state = read(self.state)
            current = self.context(state)
            self.document(current)  # Includes the cross-conversation ownership guard.
            doc = {'schemaVersion': 1, 'sessionId': self.session, 'turnId': current['turnId'],
                   'requestId': str(uuid.uuid4()), 'requirements': req, 'userQuote': quote.strip()} if req else {}
            save(self.temp, doc)
            current.update(active=True, phase='parameters' if req else 'default')
            save(self.state, state)
            return {'stored': True, 'route': 'parameters' if req else 'default_latest_10', 'requirements': req}

    def build(self, builder):
        with self.locked():
            state = read(self.state)
            current = self.context(state)
            doc = self.document(current)
            req = doc.get('requirements', {})
            result = builder(req)
            # Clear only after both fixed batch and review writes succeeded.
            save(self.temp, {})
            current.update(active=True, phase='default', lastBatchId=result['batchId'])
            save(self.state, state)
            return {**result, 'requirementsConsumed': bool(doc), 'requirementsCleared': True}

    def complete(self, batch_id):
        """Completion never destroys a newer request belonging to a later cycle/session."""
        with self.locked():
            state = read(self.state)
            for current in state.get('sessions', {}).values():
                if current.get('lastBatchId') == batch_id and current.get('phase') != 'parameters':
                    current.update(active=False, phase='completed')
            save(self.state, state)
            return {'batchId': batch_id, 'requirementsEmpty': not bool(read(self.temp)),
                    'newerRequirementsPreserved': bool(read(self.temp))}


def prompt_hook(root, payload):
    event = payload.get('hook_event_name') or ''
    if event not in ('UserPromptSubmit', 'SessionStart', 'PreInvocation'):
        return {}
    session = payload.get('session_id') or payload.get('sessionId') or payload.get('conversation_id') or payload.get('conversationId')
    if not session:
        return {}
    turn = payload.get('turn_id') or payload.get('turnId') or payload.get('invocationNum')
    req = Requirements(root, str(session), str(turn) if turn is not None else 'hook-event-' + uuid.uuid4().hex)
    with req.locked():
        state = read(req.state)
        old = state.get('sessions', {}).get(req.session, {})
        current = req.context(state)
        save(req.state, state)
        if old.get('turnId') == current['turnId'] or not current.get('active'):
            return {}
    return {'hookSpecificOutput': {'hookEventName': event, 'additionalContext':
        '【本轮需求入口】新用户回合已初始化。搜索组件开始时先调用 recruiting_application_requirements_begin；'
        '用户有本轮参数则覆盖写入需求临时JSON，再由固定脚本读取；空表走最新10家公司默认路线。'
        '同回合选择默认后不重复询问；已冻结的投递任务继续原任务，不因初始化重复组批。'}}


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--workspace', default=os.environ.get('CAMPUS_RECRUITING_WORKSPACE', '.'))
    ap.add_argument('--session')
    ap.add_argument('--turn')
    ap.add_argument('--host', choices=['codex', 'claude', 'agy'], default='codex')
    ap.add_argument('--event')
    sub = ap.add_subparsers(dest='cmd', required=True)
    sub.add_parser('begin')
    p = sub.add_parser('set')
    p.add_argument('--requirements', required=True, help='JSON对象；{} 表示默认')
    p.add_argument('--quote', required=True)
    sub.add_parser('hook')
    a = ap.parse_args()
    if a.cmd == 'hook':
        payload = json.load(sys.stdin)
        if a.event:
            payload['hook_event_name'] = a.event
        out = prompt_hook(a.workspace, payload)
        if a.host == 'agy' and out:
            out = {'injectSteps': [{'ephemeralMessage': out['hookSpecificOutput']['additionalContext']}]}
        print(json.dumps(out, ensure_ascii=False))
        return
    store = Requirements(a.workspace, a.session, a.turn)
    out = store.begin() if a.cmd == 'begin' else store.set(json.loads(a.requirements), a.quote)
    print(json.dumps(out, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
