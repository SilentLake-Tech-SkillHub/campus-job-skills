"""Read the shared host capability record; search and application use different criteria."""
import json
import os
from pathlib import Path


def host_id(env):
    explicit = env.get('RECRUITING_HOST_ID') or env.get('CAMPUS_SEARCH_HOST_ID')
    if explicit:
        return explicit
    for name, host in [('DSH_SESSION_ID', 'dsh'), ('CODEX_THREAD_ID', 'codex'), ('CODEX_SESSION_ID', 'codex'), ('CLAUDE_SESSION_ID', 'claude-code')]:
        if env.get(name):
            return host
    raise ValueError('BROWSER_HOST_ID_REQUIRED')


def select_channels(root, purpose, requested=None, env=None):
    env = os.environ if env is None else env
    if purpose not in ('search', 'application'):
        raise ValueError('INVALID_BROWSER_PURPOSE')
    file = Path(env.get('RECRUITING_BROWSER_CAPABILITY_FILE') or Path(root) / '.runtime/browser-capability.json')
    try:
        document = json.loads(file.read_text(encoding='utf-8'))
        record = document['hosts'][host_id(env)]
        if document.get('schemaVersion') != 2:
            raise KeyError('schemaVersion')
    except (OSError, ValueError, KeyError, TypeError) as exc:
        raise ValueError('BROWSER_CAPABILITY_DETECTION_REQUIRED') from exc
    iab = record.get('iab') or {}
    application = iab.get('result') == 'usable'
    search = (application or iab.get('declared') is True and iab.get('exposure') == 'agent_only'
              or iab.get('identity', {}).get('verified') is True
              and all(iab.get('actions', {}).get(k) is True for k in ('navigate', 'evaluate')))
    observed = iab.get(purpose, {}).get('usable')
    usable = (search if purpose == 'search' else application) and observed is not False
    available = (['chrome'] if document.get('chrome', {}).get('result') == 'usable' else []) + (['iab'] if usable else [])
    if not available:
        raise ValueError('NO_USABLE_BROWSER_CHANNEL')
    if requested is not None and (not isinstance(requested, list) or not requested
                                 or len(set(requested)) != len(requested)
                                 or any(ch not in ('chrome', 'iab') for ch in requested)):
        raise ValueError('INVALID_BATCH_BROWSER_CHANNELS')
    wanted = requested or available
    intersection = [ch for ch in wanted if ch in available]
    fallback = not intersection
    return {'purpose': purpose, 'requestedChannels': wanted, 'availableChannels': available,
            'channels': intersection or available, 'automaticFallback': fallback,
            'reason': 'requested_channels_unavailable_using_available_channels' if fallback else 'using_available_requested_channels'}


def application_policy_file(root, env=None):
    env = os.environ if env is None else env
    return Path(env.get('RECRUITING_PRIVATE_DIR') or Path(root) / '.runtime/application') / 'runs/batch-browser-policy.json'
