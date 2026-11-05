import json
import tempfile
import unittest
from pathlib import Path
from browser_capabilities import select_channels


def fixture_environment(root, agent_only=False):
    file = Path(root) / '.runtime/browser-capability.json'
    file.parent.mkdir(parents=True, exist_ok=True)
    iab = {'declared': True, 'result': 'agent_only' if agent_only else 'usable', 'exposure': 'agent_only' if agent_only else 'code'}
    file.write_text(json.dumps({'schemaVersion': 2, 'chrome': {'result': 'usable'}, 'hosts': {'synthetic': {'iab': iab}}}))
    return {'RECRUITING_HOST_ID': 'synthetic', 'RECRUITING_BROWSER_CAPABILITY_FILE': str(file),
            'RECRUITING_PRIVATE_DIR': str(Path(root) / '.runtime/application')}


class BrowserCapabilityTests(unittest.TestCase):
    def test_agent_browser_is_searchable_without_becoming_an_application_executor(self):
        with tempfile.TemporaryDirectory() as tmp:
            env = fixture_environment(tmp, agent_only=True)
            self.assertEqual(select_channels(tmp, 'search', env=env)['channels'], ['chrome', 'iab'])
            result = select_channels(tmp, 'application', ['iab'], env=env)
            self.assertEqual(result['channels'], ['chrome'])
            self.assertTrue(result['automaticFallback'])

    def test_missing_capability_has_no_default_dual_channel_fallback(self):
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaisesRegex(ValueError, 'DETECTION_REQUIRED'):
                select_channels(tmp, 'search', env={'RECRUITING_HOST_ID': 'synthetic'})

    def test_false_identity_cannot_become_search_capability(self):
        with tempfile.TemporaryDirectory() as tmp:
            env = fixture_environment(tmp)
            p = Path(env['RECRUITING_BROWSER_CAPABILITY_FILE'])
            d = json.loads(p.read_text());d['hosts']['synthetic']['iab'] = {'declared': True, 'result': 'declared_but_fake', 'search': {'usable': True}}
            p.write_text(json.dumps(d))
            self.assertEqual(select_channels(tmp, 'search', env=env)['channels'], ['chrome'])
