"""Fixed events invoke the real Hook in isolated subprocesses; no workbook IO."""
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

HERE = Path(__file__).resolve().parent
HOOK = HERE / 'status_sync_hook.py'
CASES = json.loads((HERE / 'fixtures/data_protection_events.json').read_text(encoding='utf-8'))


class HookProcessTests(unittest.TestCase):
    def expand(self, value, root):
        if isinstance(value, str): return value.replace('{{ROOT}}', str(root))
        if isinstance(value, list): return [self.expand(v, root) for v in value]
        if isinstance(value, dict): return {k: self.expand(v, root) for k, v in value.items()}
        return value

    def invoke(self, root, event, host):
        result = subprocess.run([sys.executable, '-B', str(HOOK), '--workspace', str(root), '--host', host],
                                input=json.dumps(event, ensure_ascii=False), text=True,
                                capture_output=True, timeout=20)
        self.assertEqual(result.returncode, 0, result.stderr)
        return json.loads(result.stdout)

    def run_case(self, case, host):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp).resolve()
            (root / '产品管理').mkdir(); (root / 'other').mkdir()
            (root / 'other/note').write_text('ordinary unrelated fixture\n')
            (root / 'book-alias').symlink_to(root / 'book.xlsx')
            script = root / '.agents/skills/campus-job-application/scripts/json_excel_sync.py'
            script.parent.mkdir(parents=True); script.write_text('# never executed by the test\n')
            cfg = {'schemaVersion': 1, 'rowKind': 'role', 'labels': [{'value': '岗位待投递'}],
                   'workbook': 'book.xlsx', 'masterData': 'data/main.json', 'confirmationLedger': 'ledger.json',
                   'companyStore': {'dir': 'company'}, 'syncScript': str(script), 'column': '投递状态', 'sheets': []}
            cfg.update(case.get('config', {}))
            (root / '产品管理/投递状态标签.json').write_text(json.dumps(cfg, ensure_ascii=False), encoding='utf-8')
            if case.get('approve'):
                self.invoke(root, {'hook_event_name': 'UserPromptSubmit', 'session_id': 'fixture',
                                   'prompt': '批准修改组件'}, host)
            event = {'hook_event_name': 'PreToolUse', 'session_id': 'fixture',
                     'cwd': self.expand(case.get('cwd', '{{ROOT}}'), root)}
            args = self.expand(case['args'], root)
            if case.get('shape') == 'dsh': event['toolCall'] = {'name': case['tool'], 'arguments': args}
            elif case.get('shape') == 'agy': event.update(toolName=case['tool'], toolInput=args)
            elif case.get('shape') == 'serialized': event['toolCall'] = {'name': case['tool'], 'arguments': json.dumps(args)}
            elif case.get('shape') == 'raw': event['toolCall'] = {'name': case['tool'], 'arguments': args}
            else: event.update(tool_name=case['tool'], tool_input=args)
            result = self.invoke(root, event, host)
            denied = result.get('decision') == 'deny' or result.get('hookSpecificOutput', {}).get('permissionDecision') == 'deny'
            reason = result.get('reason') or result.get('hookSpecificOutput', {}).get('permissionDecisionReason', '')
            self.assertEqual(not denied, case['allow'], f'{case["name"]}: {reason}')
            if denied and not case['name'].startswith('soft-lock'):
                self.assertIn('数据文件保护', reason)
                self.assertIn(str(HOOK), reason)
                self.assertIn('正确做法', reason)
            self.assertFalse((root / cfg['workbook']).exists())

    def test_fixed_events_all_hosts(self):
        for host in ['codex', 'claude', 'agy']:
            for case in CASES:
                with self.subTest(host=host, event=case['name']): self.run_case(case, host)


class SharedToolArgumentsTests(unittest.TestCase):
    """The text-tool list and file argument names come only from the application component's file."""

    def test_definition_is_the_application_component_file(self):
        import hook_file_policy as policy
        import status_sync_hook as hook
        spec = json.loads(policy.TOOL_ARGUMENTS_FILE.read_text(encoding='utf-8'))
        self.assertEqual(policy.TOOL_ARGUMENTS_FILE.parent.parent.name, 'job-application-assistant-v2')
        self.assertEqual(set(policy.TEXT_TOOLS), {t.lower() for t in spec['textTools']})
        self.assertEqual(policy.PATH_KEYS, tuple(spec['fileArgumentKeys']))
        self.assertIs(hook.PATH_KEYS, policy.PATH_KEYS)
        self.assertFalse(hasattr(hook, 'TEXT_TOOLS'))

    def test_file_arguments_at_any_depth_make_a_text_tool_a_file_operation(self):
        import hook_file_policy as policy
        self.assertTrue(policy.text_only('todo_write', {'todos': [{'content': 'check book.xlsx'}]}))
        for args in ({'todos': [], 'TargetFile': 'book.xlsx'}, {'message': {'attachments': ['book.xlsx']}},
                     {'questions': [{'question': 'q', 'notebook_path': 'book.xlsx'}]}, ['book.xlsx']):
            self.assertFalse(policy.text_only('send_message', args), args)

    def test_missing_definition_trusts_no_tool_as_text(self):
        import hook_file_policy as policy
        self.assertEqual(policy.tool_arguments(Path(tempfile.gettempdir()) / 'no-such-tool-arguments.json'), (frozenset(), ()))


if __name__ == '__main__': unittest.main(verbosity=2)
