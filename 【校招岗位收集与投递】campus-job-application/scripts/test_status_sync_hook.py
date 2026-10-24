import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import status_sync_hook as hook
import json_excel_sync as store
import workbook_store
from openpyxl import Workbook


class HookDetailsTests(unittest.TestCase):
    def test_actual_before_after_for_status_and_other_fields(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / '产品管理').mkdir()
            wb = Workbook()
            ws = wb.active
            ws.title = '中国大陆'
            ws.append(['公司', '岗位名称', '投递状态', '备注'])
            ws.append(['甲公司', 'AI产品经理', '岗位待筛选', '原备注'])
            wb.save(root / 'book.xlsx')
            master = workbook_store.snapshot(root / 'book.xlsx')
            row = master['sheets'][0]['rows'][1]
            store.set_cell(row, 3, '岗位待投递')
            store.set_cell(row, 4, '用户已确认，第一优先')
            (root / 'master.json').write_text(json.dumps(master), encoding='utf-8')
            cfg = {'schemaVersion': 1, 'rowKind': 'role', 'column': '投递状态', 'sheets': ['中国大陆'],
                   'labels': [{'value': '岗位待筛选'}, {'value': '岗位待投递'}], 'masterData': 'master.json', 'workbook': 'book.xlsx',
                   'confirmationLedger': 'ledger.json'}
            (root / '产品管理/投递状态标签.json').write_text(json.dumps(cfg), encoding='utf-8')
            before = (root / 'book.xlsx').read_bytes()
            result = hook.pending_details(root)
            self.assertIn('实际变化：新增 0 行、删除 0 行、修改 1 行（2 个单元格）', result)
            self.assertIn('当前「岗位待筛选」→写入「岗位待投递」', result)
            self.assertIn('甲公司｜AI产品经理', result)
            self.assertIn('原备注', result)
            self.assertIn('用户已确认，第一优先', result)
            self.assertEqual(before, (root / 'book.xlsx').read_bytes())

    def test_pending_message_requires_details_before_question(self):
        with patch.object(hook, 'sync', return_value={'unwritten': True}), \
             patch.object(hook, 'pending_details', return_value='甲公司：旧值→新值'):
            text, _ = hook.pending_text(Path('.'))
        self.assertIn('甲公司：旧值→新值', text)
        self.assertIn('禁止只报总数', text)

    def test_patch_guard_checks_target_not_body(self):
        cfg = {'workbook': 'book.xlsx', 'masterData': 'master.json', 'confirmationLedger': 'ledger.json'}
        body = '*** Begin Patch\n*** Update File: /work/status_sync_hook.py\n+read master.json\n*** End Patch'
        self.assertIsNone(hook.tool_violation(cfg, {'tool_name': 'apply_patch', 'tool_input': body}))
        for action in ['Update', 'Add', 'Delete']:
            body = f'*** Begin Patch\n*** {action} File: /work/master.json\n*** End Patch'
            self.assertIn('master.json', hook.tool_violation(cfg, {'tool_name': 'apply_patch', 'tool_input': body}))
        body = '*** Begin Patch\n*** Update File: /work/safe.py\n*** Move to: /work/master.json\n*** End Patch'
        self.assertIn('master.json', hook.tool_violation(cfg, {'tool_name': 'apply_patch', 'tool_input': body}))

    def test_patch_command_envelope_still_checks_actual_target(self):
        cfg = {'workbook': 'book.xlsx', 'masterData': 'master.json', 'confirmationLedger': 'ledger.json'}
        safe = '*** Begin Patch\n*** Update File: /work/test_example.py\n+writeFile master.json\n*** End Patch'
        self.assertIsNone(hook.tool_violation(cfg, {'tool_name': 'apply_patch', 'tool_input': {'command': safe}}))
        wrapped = 'text(await tools.apply_patch(' + json.dumps(safe) + '));'
        self.assertIsNone(hook.tool_violation(cfg, {'tool_name': 'functions.exec', 'tool_input': {'input': wrapped}}))
        denied = safe.replace('/work/test_example.py', '/work/master.json')
        for name, args in [('apply_patch', {'command': denied}), ('functions.exec', {'input': 'text(await tools.apply_patch(' + json.dumps(denied) + '));'})]:
            self.assertIn('master.json', hook.tool_violation(cfg, {'tool_name': name, 'tool_input': args}))
        self.assertIsNotNone(hook.tool_violation(cfg, {'tool_name': 'functions.exec', 'tool_input': {'input': wrapped + '; write_text("master.json")'}}))


class LabelGroupTests(unittest.TestCase):
    def test_rules_text_and_guard_cover_every_label_group(self):
        """Hook 与脚本同一入口加载全部标签组：注入文本按组列出，每个组文件都受写入保护。"""
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / '产品管理').mkdir()
            cfg = {'schemaVersion': 1, 'rowKind': 'role', 'labelGroups': ['产品管理/公司状态标签.json'],
                   'labels': [{'value': '岗位待筛选', 'definition': '待用户筛选'}],
                   'workbook': 'book.xlsx', 'masterData': 'master.json', 'confirmationLedger': 'ledger.json'}
            (root / '产品管理/投递状态标签.json').write_text(json.dumps(cfg, ensure_ascii=False), encoding='utf-8')
            (root / '产品管理/公司状态标签.json').write_text(json.dumps(
                {'schemaVersion': 1, 'rowKind': 'company', 'labels': [{'value': '已有岗位', 'definition': '有岗位行'}]},
                ensure_ascii=False), encoding='utf-8')
            text = hook.rules_text(cfg, root)
            self.assertIn('岗位行只能用（产品管理/投递状态标签.json）：「岗位待筛选」=待用户筛选', text)
            self.assertIn('公司行只能用（产品管理/公司状态标签.json）：「已有岗位」=有岗位行', text)
            self.assertIn('产品管理/公司状态标签.json', hook.protected(cfg))
            body = '*** Begin Patch\n*** Update File: /work/产品管理/公司状态标签.json\n*** End Patch'
            self.assertIsNotNone(hook.tool_violation(cfg, {'tool_name': 'apply_patch', 'tool_input': body}))


class BrowserRuleTests(unittest.TestCase):
    def test_browser_section_is_injected_from_the_single_config(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / '产品管理').mkdir()
            self.assertIn('读取 产品管理/求职配置.md 失败', hook.browser_rules(root))
            (root / '产品管理/求职配置.md').write_text('# 配置\n\n## 浏览器\n\n- 默认双开。\n- 单开用内置浏览器。\n\n## 组件\n\n- x\n', encoding='utf-8')
            text = hook.browser_rules(root)
            self.assertIn('唯一来源', text)
            self.assertIn('- 默认双开。\n- 单开用内置浏览器。', text)
            self.assertNotIn('组件', text.split('】', 1)[1])
            (root / '产品管理/求职配置.md').write_text('# 配置\n\n## 组件\n', encoding='utf-8')
            self.assertIn('缺少「## 浏览器」', hook.browser_rules(root))


class StackGuardTests(unittest.TestCase):
    def check(self, tool, ti):
        return hook.stack_violation(Path('/work'), {'tool_name': tool, 'tool_input': ti})

    def test_only_fixed_files_under_handoff_and_no_copies(self):
        ok = [('Write', {'file_path': '/work/岗位交接/公司/中国大陆/甲公司/公司.json'}),
              ('Edit', {'file_path': '/work/岗位交接/投递批次.json'}),
              ('Write', {'file_path': '/work/岗位交接/公司/中国大陆/网易/登录@leihuo.163.com.json'}),
              ('Write', {'file_path': '/work/岗位交接/公司/中国大陆/网易/表单定义.json'}),
              ('Bash', {'command': 'python3 company_store.py refresh'}),
              ('Bash', {'command': 'cat 岗位交接/投递批次.json'}),
              ('Bash', {'command': 'rm -rf 岗位交接/2026-10-18-旧批次'})]
        for tool, ti in ok:
            self.assertIsNone(self.check(tool, ti), (tool, ti))
        bad = [('Write', {'file_path': '/work/岗位交接/2026-10-21-新批次/公司方案.json'}),
               ('Write', {'file_path': '/work/岗位交接/公司/中国大陆/甲公司/研究记录.json'}),
               ('Write', {'file_path': '/work/岗位交接/公司/中国大陆/甲公司/表单/模板.json'}),
               ('Write', {'file_path': '/work/产品管理/求职目标-上一版.json'}),
               ('Write', {'file_path': '/work/简历/历史/a.md'}),
               ('apply_patch', '*** Begin Patch\n*** Add File: /work/岗位交接/备份/x.json\n*** End Patch'),
               ('Bash', {'command': 'cp 岗位交接/工作簿全量.json 岗位交接/工作簿全量-副本.json'}),
               ('Bash', {'command': "python3 -c \"open('岗位交接/草稿.json','w').write('x')\""})]
        for tool, ti in bad:
            self.assertIn('记录唯一门禁', self.check(tool, ti) or '', (tool, ti))



class ComponentLockTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name).resolve()
        (self.root / '产品管理').mkdir()

    def tearDown(self):
        self.tmp.cleanup()

    def check(self, tool, ti, cwd=None):
        payload = {'tool_name': tool, 'tool_input': ti}
        if cwd:
            payload['cwd'] = str(self.root / cwd)
        return hook.component_violation(self.root, payload)

    def test_components_locked_until_user_enables_maintenance(self):
        r = self.root
        v2 = '.agents/skills/job-application-assistant-v2'
        cj = '.agents/skills/campus-job-application'
        ok = [('Bash', {'command': f'python3 {cj}/scripts/json_excel_sync.py check > /tmp/out.json 2>&1'}, None),
              ('Bash', {'command': f'cd {v2} && node --test test/ 2>&1 | tail -5'}, None),
              ('Bash', {'command': f'cat {v2}/shared/role-queue.mjs'}, None),
              ('Bash', {'command': f'cp {cj}/scripts/json_excel_sync.py /tmp/x.py'}, None),
              ('Bash', {'command': f'git -C {v2} log --oneline -3'}, None),
              ('Write', {'file_path': str(r / '产品管理/求职目标.json')}, None),
              ('Read', {'file_path': str(r / v2 / 'SKILL.md')}, None)]
        for tool, ti, cwd in ok:
            self.assertIsNone(self.check(tool, ti, cwd), (tool, ti))
        bad = [('Edit', {'file_path': str(r / cj / 'scripts/json_excel_sync.py')}, None),
               ('Write', {'file_path': f'{v2}/shared/x.mjs'}, None),
               ('apply_patch', f'*** Begin Patch\n*** Update File: {r}/{v2}/SKILL.md\n*** End Patch', None),
               ('Bash', {'command': f"sed -i 's/a/b/' {cj}/SKILL.md"}, None),
               ('Bash', {'command': "sed -i 's/a/b/' shared/x.mjs"}, v2),
               ('Bash', {'command': f'cd {v2} && echo x > shared/y.mjs'}, None),
               ('Bash', {'command': f'cp /tmp/x.py {cj}/scripts/company_store.py'}, None),
               ('Bash', {'command': f'git -C {v2} checkout -- shared'}, None),
               ('Bash', {'command': 'npm install left-pad'}, v2),
               ('Bash', {'command': f"python3 -c \"open('{cj}/scripts/a.py','w').write('x')\""}, None),
               ('Edit', {'file_path': str(r / '.claude/settings.json')}, None),
               ('Bash', {'command': f'rm -rf {v2}/test'}, None)]
        for tool, ti, cwd in bad:
            self.assertIn('组件只读', self.check(tool, ti, cwd) or '', (tool, ti))
        # 开关文件 AI 永远不能写
        for tool, ti in [('Write', {'file_path': str(r / hook.SWITCH)}), ('Bash', {'command': f'echo {{}} > {hook.SWITCH}'})]:
            self.assertIn('只能由用户本人修改', self.check(tool, ti) or '')
        # 用户开启维护后放开组件，开关文件仍不可写
        (r / hook.SWITCH).write_text('{"开启": true}', encoding='utf-8')
        self.assertIsNone(self.check('Edit', {'file_path': str(r / cj / 'scripts/json_excel_sync.py')}))
        self.assertIn('只能由用户本人修改', self.check('Write', {'file_path': str(r / hook.SWITCH)}) or '')
        (r / hook.SWITCH).write_text('{"开启": false}', encoding='utf-8')
        self.assertIn('组件只读', self.check('Edit', {'file_path': str(r / cj / 'scripts/json_excel_sync.py')}) or '')

    def test_user_approval_in_chat_unlocks_only_that_session(self):
        r, cj = self.root, '.agents/skills/campus-job-application'
        edit = lambda sid: hook.component_violation(r, {'tool_name': 'Edit', 'session_id': sid,
                                                       'tool_input': {'file_path': str(r / cj / 'scripts/json_excel_sync.py')}})
        say = lambda sid, text: hook.record_approval(r, {'hook_event_name': 'UserPromptSubmit', 'session_id': sid, 'prompt': text})
        self.assertIn('组件只读', edit('s1') or '')
        for text in ('组件可以改吗？', '不要修改组件', '先别批准改 Hook', '这个组件锁做成软锁，允许在用户批准的情况下修改'):
            say('s1', text)
            self.assertIn('组件只读', edit('s1') or '', text)
        self.assertIn('已记录用户批准', say('s3', '批准修改组件'))
        self.assertIn('已记录用户批准', say('s4', '授权你改 Hook'))
        self.assertIn('已记录用户批准', say('s1', '内部组件本次授权给你修改'))
        self.assertIsNone(edit('s1'))
        self.assertIn('组件只读', edit('s2') or '')  # 只放开批准所在的会话
        self.assertIn('批准记录', hook.component_violation(r, {'tool_name': 'Write', 'session_id': 's1',
                                                         'tool_input': {'file_path': str(r / hook.APPROVALS)}}) or '')
        self.assertIn('只能由用户本人修改', hook.component_violation(r, {'tool_name': 'Write', 'session_id': 's1',
                                                              'tool_input': {'file_path': str(r / hook.SWITCH)}}) or '')
        self.assertIn('收回', say('s1', '收回授权'))
        self.assertIn('组件只读', edit('s1') or '')

class ApprovalAnswerParsingTests(unittest.TestCase):
    def envelope(self, answers, question='是否允许修改组件？不要修改旧数据'):
        items = [{'questionItemId': json.dumps(['request_user_input_async', 'call_test', i]),
                  'question': question, 'answer': answer} for i, answer in enumerate(answers)]
        return '<send_user_message_question_reply>\n' + json.dumps(items) + '\n</send_user_message_question_reply>'

    def test_only_actual_answers_are_scanned(self):
        prompt = self.envelope(['应届毕业生', '这5项留空', '批准修改组件'])
        self.assertEqual(hook._approval_prompts({'prompt': prompt}), ['应届毕业生', '这5项留空', '批准修改组件'])
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            payload = {'session_id': 'owner', 'prompt': prompt}
            self.assertIn('已记录用户批准', hook.record_approval(root, payload))
            self.assertTrue(hook.approved(root, payload))
            self.assertFalse(hook.approved(root, {'session_id': 'other'}))
            quote = json.loads((root / hook.APPROVALS).read_text())['owner']['userQuote']
            self.assertEqual(quote, '批准修改组件')
            hook.record_approval(root, {'session_id': 'owner', 'prompt': self.envelope(['批准修改组件', '收回授权'])})
            self.assertFalse(hook.approved(root, payload))

    def test_questions_and_nested_tool_text_never_authorize(self):
        for answer in ['不要修改组件', '是否批准修改组件', '', '否']:
            with tempfile.TemporaryDirectory() as tmp:
                root = Path(tmp)
                payload = {'session_id': 's', 'prompt': self.envelope([answer], question='批准修改组件')}
                hook.record_approval(root, payload)
                self.assertFalse(hook.approved(root, payload))
        self.assertEqual(hook._approval_prompts({'tool_result': {'content': '批准修改组件'}}), [])
        self.assertEqual(hook._approval_prompts({'prompt': '批准修改组件'}), ['批准修改组件'])

    def test_malformed_envelopes_fail_closed(self):
        valid = self.envelope(['批准修改组件'])
        bad = [valid + '批准修改组件', '批准修改组件' + valid, valid.replace('call_test', 'bad'),
               valid.replace('request_user_input_async', 'tool_result'),
               valid.replace('"answer":', '"extra": 1, "answer":'),
               valid.replace('"answer":', '"answer": "批准修改组件", "answer":'),
               '<send_user_message_question_reply>{"answer":"批准修改组件"}</send_user_message_question_reply>']
        for index in [-1, True, 0.5]:
            bad.append(valid.replace('0]', json.dumps(index) + ']'))
        for prompt in bad:
            self.assertEqual(hook._approval_prompts({'prompt': prompt}), [], prompt)


class ConversationReminderTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        folder = self.root / '.codex/hooks'
        folder.mkdir(parents=True)
        (folder / 'scope.json').write_text(json.dumps({'sessions': {
            'application-session': {'status_sync_reminders': False},
            'search-session': {'workflow': 'search'},
        }}), encoding='utf-8')

    def invoke(self, event, session='application-session', **extra):
        import io
        payload = {'hook_event_name': event, 'session_id': session, **extra}
        cfg = {'workbook': 'book.xlsx', 'masterData': 'master.json',
               'confirmationLedger': 'ledger.json'}
        out = io.StringIO()
        with patch.object(hook.sys, 'argv', ['hook', '--workspace', str(self.root), '--host', 'codex']), \
             patch.object(hook.sys, 'stdin', io.StringIO(json.dumps(payload))), \
             patch.object(hook.sys, 'stdout', out), \
             patch.object(hook, 'load_cfg', return_value=cfg), \
             patch.object(hook, 'record_approval', return_value='') as approval, \
             patch.object(hook, 'rules_text', return_value='RULES'), \
             patch.object(hook, 'browser_rules', return_value='BROWSER'), \
             patch.object(hook, 'pending_text', return_value=('PENDING', {})) as pending, \
             patch.object(hook, 'sync', return_value={}) as sync:
            hook.main()
            return json.loads(out.getvalue()), pending.called, sync.called, approval.called

    def test_disabled_session_skips_only_excel_reminders(self):
        """关闭写入提醒的会话：不提示、不拦截写入 Excel，但门禁规则和浏览器规则照常注入。"""
        for event in ('SessionStart', 'UserPromptSubmit', 'SubagentStart', 'PreInvocation'):
            result, pending, sync, _ = self.invoke(event)
            self.assertEqual(result['hookSpecificOutput']['additionalContext'], 'RULESBROWSER', event)
            self.assertFalse(pending, event)
        for event in ('Stop', 'SubagentStop'):
            result, pending, sync, _ = self.invoke(event)
            self.assertEqual(result, {}, event)
            self.assertFalse(pending, event)
            self.assertFalse(sync, event)

    def test_other_session_keeps_startup_reminder_and_stop_block(self):
        result, pending, _, _ = self.invoke('UserPromptSubmit', 'search-session')
        self.assertEqual(result['hookSpecificOutput']['additionalContext'], 'RULESBROWSERPENDING')
        self.assertTrue(pending)
        result, pending, sync, _ = self.invoke('Stop', 'search-session')
        self.assertEqual(result['decision'], 'block')
        self.assertTrue(pending)
        self.assertTrue(sync)

    def test_disabled_session_keeps_write_protection(self):
        result, _, _, _ = self.invoke('PreToolUse', tool_name='Edit',
                                      tool_input={'file_path': str(self.root / 'master.json')})
        self.assertEqual(result['hookSpecificOutput']['permissionDecision'], 'deny')

    def test_soft_lock_approval_processing_remains_active(self):
        _, _, _, approval = self.invoke('UserPromptSubmit')
        self.assertTrue(approval)

    def test_explicit_payload_identity_wins_over_host_environment(self):
        with patch.dict(hook.os.environ, {'CODEX_THREAD_ID': 'application-session'}):
            self.assertFalse(hook.sync_reminders_enabled(self.root, {}))
            self.assertTrue(hook.sync_reminders_enabled(self.root, {'session_id': 'search-session'}))


if __name__ == '__main__':
    unittest.main()
