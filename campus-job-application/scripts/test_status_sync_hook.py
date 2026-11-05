import json
import os
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

    def test_data_gate_uses_the_same_write_detection_as_the_component_lock(self):
        cfg = {'workbook': 'book.xlsx', 'masterData': '岗位交接/工作簿全量.json', 'confirmationLedger': 'ledger.json',
               'companyStore': {'dir': '岗位交接/公司'}}
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp).resolve()
            f = '岗位交接/公司/中国大陆/Keep/登录.json'
            writes = [f"node -e \"require('fs').writeFileSync('{f}','{{}}')\"",
                      f"node -e \"require('fs').renameSync('/tmp/x.json','{f}')\"",
                      f"python3 -c \"import os; os.replace('/tmp/x.json','{f}')\"",
                      f"perl -pi -e 's/a/b/' {f}",
                      f"echo '{{}}' > {f}",
                      f"cp /tmp/x.json {f}"]
            for cmd in writes:
                self.assertIn('岗位交接/公司/', hook.tool_violation(cfg, {'tool_name': 'Bash', 'tool_input': {'command': cmd}, 'cwd': str(root)}, root) or [], cmd)
            for cmd in [f'cat {f}', f"node -e \"console.log(require('fs').readFileSync('{f}','utf8'))\"",
                        'node .agents/skills/job-application-assistant-v2/scripts/site-entry.mjs --workspace . --site S0015 --input /tmp/keep-entry.json']:
                self.assertIsNone(hook.tool_violation(cfg, {'tool_name': 'Bash', 'tool_input': {'command': cmd}, 'cwd': str(root)}, root), cmd)

    def test_patch_guard_compares_full_path_not_file_name(self):
        cfg = {'workbook': 'book.xlsx', 'masterData': '岗位交接/master.json', 'confirmationLedger': 'ledger.json'}
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            other = '*** Begin Patch\n*** Update File: ' + str(root / 'tests/fixtures/master.json') + '\n*** End Patch'
            self.assertIsNone(hook.tool_violation(cfg, {'tool_name': 'apply_patch', 'tool_input': other}, root))
            real = '*** Begin Patch\n*** Update File: ' + str(root / '岗位交接/master.json') + '\n*** End Patch'
            self.assertIn('岗位交接/master.json', hook.tool_violation(cfg, {'tool_name': 'apply_patch', 'tool_input': real}, root))
            rel = '*** Begin Patch\n*** Update File: master.json\n*** End Patch'
            self.assertIn('岗位交接/master.json', hook.tool_violation(cfg, {'tool_name': 'apply_patch', 'tool_input': rel, 'cwd': str(root / '岗位交接')}, root))
            self.assertIsNone(hook.tool_violation(cfg, {'tool_name': 'apply_patch', 'tool_input': rel, 'cwd': str(root / 'tests')}, root))


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
        ok = [('Write', {'file_path': '/work/岗位交接/公司/C00001.json'}),
              ('Edit', {'file_path': '/work/岗位交接/投递回传.json'}),
              ('Write', {'file_path': '/work/岗位交接/公司/中国大陆/网易/登录@leihuo.163.com.json'}),
              ('Write', {'file_path': '/work/岗位交接/公司/中国大陆/网易/表单定义.json'}),
              ('Bash', {'command': 'python3 company_store.py refresh'}),
              ('Bash', {'command': 'cat 岗位交接/公司/C00001.json'}),
              ('Bash', {'command': 'rm -rf 岗位交接/2026-10-18-旧批次'})]
        for tool, ti in ok:
            self.assertIsNone(self.check(tool, ti), (tool, ti))
        bad = [('Write', {'file_path': '/work/岗位交接/2026-10-21-新批次/公司方案.json'}),
               ('Write', {'file_path': '/work/岗位交接/公司/中国大陆/甲公司/公司.json'}),
               ('Write', {'file_path': '/work/岗位交接/投递批次.json'}),
               ('Write', {'file_path': '/work/岗位交接/公司/C1.json'}),
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


    def test_unconditional_approval_is_not_denied_by_an_independent_browser_question(self):
        accepted = [
            '批准修改组件。是否有内置 browser？',
            '批准修改组件，另外请确认是否有内置 browser？',
            '批准修改组件\n是否有内置 browser？',
            '授权你改 Hook；另外是否支持内置浏览器？',
            '批准修改组件，不要修改个人资料。',
        ]
        for index, prompt in enumerate(accepted):
            with tempfile.TemporaryDirectory() as tmp:
                root = Path(tmp)
                payload = {'session_id': 'approved', 'prompt': prompt}
                self.assertIn('已记录用户批准', hook.record_approval(root, payload), prompt)
                self.assertTrue(hook.approved(root, payload))
                data = json.loads((root / hook.APPROVALS).read_text())['approved']
                self.assertNotIn('是否', data['userQuote'])
                self.assertEqual(data['expiresAt'] - data['approvedAt'], 12 * 3600)
                self.assertFalse(hook.approved(root, {'session_id': 'other'}))

    def test_conditions_questions_and_denials_cannot_be_split_into_permission(self):
        denied = [
            '如果测试通过，批准修改组件',
            '若审批完成，批准修改组件',
            '批准修改组件。只有测试通过，批准修改组件。',
            '如果测试通过\n批准修改组件',
            '如果测试通过。批准修改组件',
            '只有确认方案以后，批准修改组件',
            '批准修改组件，前提是测试通过',
            '是否批准修改组件？',
            '批准修改组件吗？',
            '批准修改组件？',
            '不要批准修改组件。是否有内置 browser？',
            '允许在用户批准的情况下修改组件。',
        ]
        for prompt in denied:
            with tempfile.TemporaryDirectory() as tmp:
                root = Path(tmp)
                payload = {'session_id': 'blocked', 'prompt': prompt}
                hook.record_approval(root, payload)
                self.assertFalse(hook.approved(root, payload), prompt)

    def test_answer_envelope_still_uses_only_actual_answer_for_clause_approval(self):
        prompt = self.envelope(['批准修改组件。是否有内置 browser？'],
                               question='不要修改组件；是否批准修改组件？')
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            self.assertIn('已记录用户批准', hook.record_approval(root, {'session_id': 'owner', 'prompt': prompt}))
            quote = json.loads((root / hook.APPROVALS).read_text())['owner']['userQuote']
            self.assertEqual(quote, '批准修改组件')
            hook.record_approval(root, {'session_id': 'owner', 'prompt': self.envelope(['批准修改组件。是否有内置 browser？', '收回授权'])})
            self.assertFalse(hook.approved(root, {'session_id': 'owner'}))

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

    def invoke(self, event, session='application-session', pending_value=('PENDING', {}), sync_value=None, **extra):
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
             patch.object(hook, 'pending_text', return_value=pending_value) as pending, \
             patch.object(hook, 'sync', return_value=sync_value or {}) as sync:
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

    def test_stop_before_code_migration_does_not_raise_false_alarms(self):
        # 编码迁移前 check 只回 migration_required：不拦截、不问 Excel
        result, _, _, _ = self.invoke('Stop', 'search-session', pending_value=('', {}), sync_value={'migration_required': True})
        self.assertEqual(result, {})
        # 只有名额收口时只要求展示收口变更，不问“JSON 有未写入 Excel 的改动”
        result, _, _, _ = self.invoke('Stop', 'search-session', pending_value=('', {}), sync_value={'quota_full_but_unclosed': ['X：当前状态是A，更新状态是B']})
        self.assertEqual(result['decision'], 'block')
        self.assertNotIn('未写入 Excel', result['reason'])
        self.assertIn('propose', result['reason'])

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

class QuotedMaintenanceAuthorityTests(unittest.TestCase):
    def test_quoted_forwarded_and_example_approvals_are_not_authority(self):
        examples = [
            '下一步：新开一个对话，并在那个对话里说一句“批准修改组件”。',
            '他说：批准修改组件。',
            '请回复批准修改组件。',
            '原话是：授权你改 Hook。',
            '> 批准修改组件\n这里只是在引用。',
            '```text\n批准修改组件\n```',
            '`批准修改组件`',
            '“批准修改组件”',
            '```text\n批准修改组件',
            '> 引用内容\n批准修改组件\n\n这里只是说明。',
        ]
        for prompt in examples:
            with self.subTest(prompt=prompt):
                self.assertEqual(hook.component_approval_quotes(prompt), [])

    def test_direct_approval_survives_unrelated_quoted_text(self):
        self.assertEqual(hook.component_approval_quotes('批准按上述范围修改组件，包含授权误判修复。'), ['批准按上述范围修改组件'])
        self.assertEqual(hook.component_approval_quotes('批准修改组件。文档例句是“不要修改组件”。'), ['批准修改组件'])

    def test_quoted_revocation_never_cancels_a_direct_approval(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            payload = {'session_id': 'current', 'prompt': '批准修改组件。说明里的“收回授权”只是示例。'}
            self.assertIn('已记录用户批准', hook.record_approval(root, payload))
            self.assertTrue(hook.approved(root, payload))
            hook.record_approval(root, {'session_id': 'current', 'prompt': '收回授权'})
            self.assertFalse(hook.approved(root, payload))

class RevocationQuestionTests(unittest.TestCase):
    def test_questions_conditions_and_reported_revocations_do_not_revoke(self):
        for prompt in ['收回授权？','是否收回授权','如果测试失败，收回授权','他说：收回授权','不要收回授权']:
            with self.subTest(prompt=prompt):self.assertFalse(hook.direct_revocation(prompt))
        self.assertTrue(hook.direct_revocation('收回授权。'))

class ExpandedComponentLockTests(unittest.TestCase):
    def test_plan_and_preference_skills_require_the_same_session_approval(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            for name in ('campus-job-plan-inventory', 'campus-job-preference-intake'):
                event = {'session_id':'owner','tool_name':'Write','tool_input':{'file_path':str(root/'.agents/skills'/name/'SKILL.md'),'content':'test'},'cwd':str(root)}
                self.assertIn('组件只读', hook.component_violation(root,event))
            hook.record_approval(root, {'session_id':'owner','prompt':'批准修改组件'})
            self.assertIsNone(hook.component_violation(root,event))


class InjectedRuleTextIsNotApprovalTests(unittest.TestCase):
    """DSH joins the queued SessionStart context into the `prompt` of the user's first message."""
    RULE = '用户在对话里明确批准修改组件后，本会话可以修改（12 小时内有效，用户说“收回授权”即失效）。'

    def workspace(self, tmp):
        root = Path(tmp).resolve()
        (root / '产品管理').mkdir()
        cfg = {'schemaVersion': 1, 'rowKind': 'role', 'labels': [{'value': '岗位待投递', 'definition': '待投递'}],
               'workbook': 'book.xlsx', 'masterData': 'main.json', 'confirmationLedger': 'ledger.json',
               'column': '投递状态', 'sheets': []}
        (root / hook.LABELS_FILE).write_text(json.dumps(cfg, ensure_ascii=False), encoding='utf-8')
        return root, hook.rules_text(cfg, root) + hook.browser_rules(root)

    def test_rule_wording_alone_is_not_an_approval(self):
        self.assertEqual(hook.component_approval_quotes(self.RULE), [])
        self.assertEqual(hook.component_approval_quotes('用户在对话里明确批准修改组件后'), [])
        self.assertEqual(hook.component_approval_quotes('批准修改组件后'), [])
        self.assertEqual(hook.component_approval_quotes('须用户批准修改组件'), [])

    def test_injected_rules_joined_with_a_plain_message_record_nothing(self):
        with tempfile.TemporaryDirectory() as tmp:
            root, injected = self.workspace(tmp)
            for prompt in (injected, injected + '\n继续', '继续\n' + injected, injected + '\n\n继续'):
                with self.subTest(prompt=prompt[-20:]):
                    payload = {'session_id': 'dsh', 'prompt': prompt}
                    self.assertNotIn('已记录用户批准', hook.record_approval(root, payload))
                    self.assertFalse(hook.approved(root, payload))

    def test_rules_changed_after_injection_still_do_not_approve(self):
        with tempfile.TemporaryDirectory() as tmp:
            root, injected = self.workspace(tmp)
            payload = {'session_id': 'dsh', 'prompt': injected.replace('12 小时', '24 小时') + '\n继续'}
            self.assertNotIn('已记录用户批准', hook.record_approval(root, payload))
            self.assertFalse(hook.approved(root, payload))

    def test_real_words_next_to_injected_rules_still_count(self):
        with tempfile.TemporaryDirectory() as tmp:
            root, injected = self.workspace(tmp)
            payload = {'session_id': 'dsh', 'prompt': injected + '\n批准修改组件'}
            self.assertIn('已记录用户批准', hook.record_approval(root, payload))
            quote = json.loads((root / hook.APPROVALS).read_text(encoding='utf-8'))['dsh']['userQuote']
            self.assertEqual(quote, '批准修改组件')
            # The rule's own mention of revoking never cancels; the user's own words do.
            hook.record_approval(root, {'session_id': 'dsh', 'prompt': injected + '\n继续'})
            self.assertTrue(hook.approved(root, payload))
            hook.record_approval(root, {'session_id': 'dsh', 'prompt': injected + '\n收回授权'})
            self.assertFalse(hook.approved(root, payload))

    def test_explicit_approvals_are_recognized_as_before(self):
        for prompt, quote in [('批准修改组件', '批准修改组件'), ('授权你改 Hook', '授权你改 Hook'),
                              ('内部组件本次授权给你修改', '内部组件本次授权给你修改'),
                              ('我作为用户批准修改组件', '我作为用户批准修改组件'),
                              ('看完方案后，批准修改组件', '批准修改组件')]:
            with self.subTest(prompt=prompt):
                self.assertEqual(hook.component_approval_quotes(prompt), [quote])


class ReadOnlyCommandsAreNotWritesTests(unittest.TestCase):
    V2 = '.agents/skills/job-application-assistant-v2'
    SP = '/private/tmp/scratch-not-in-workspace'

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name).resolve()
        (self.root / '产品管理').mkdir()

    def tearDown(self):
        self.tmp.cleanup()

    def verdict(self, command, cwd=None):
        return hook.component_violation(self.root, {
            'tool_name': 'Bash', 'tool_input': {'command': command}, 'session_id': 'locked',
            'cwd': str(self.root / cwd) if cwd else str(self.root)})

    def test_commands_that_were_really_blocked_now_pass(self):
        v2 = self.V2
        for command in [
            # The three commands the lock refused in practice.
            f'cd "{self.root}/{v2}" && SP={self.SP}; (time node --test test/*.test.mjs) > "$SP/baseline-node.txt" 2>&1; tail -14 "$SP/baseline-node.txt"; git status --short',
            f"cd {v2} && grep -n ensureGroups shared/browser.mjs | cut -c1-120; awk '/async ensureGroups/{{f=1}} f{{print; n++}} n>=30{{exit}}' shared/browser.mjs | cut -c1-260",
            f"cd {v2} && python3 - <<'EOF'\nimport json\nfor i in range(9, -1, -1):\n    if i>=0: print(i, \"don't\")\nEOF",
        ]:
            with self.subTest(command=command[:60]):
                self.assertIsNone(self.verdict(command))

    def test_comparisons_and_text_inside_the_component_directory_pass(self):
        for command in [
            "node -e 'for(let i=9;i>=0;i--)console.log(i)'",
            "node -e 'const f = x => x.length; if (a > b.length) f(a)'",
            'python3 -c "print([x for x in rows if x > 3])"',
            "grep -n 'n>=30' SKILL.md",
            "printf 'a>b' SKILL.md",
            "grep -n 'foo > bar.txt' SKILL.md",
            "awk '$1>$2' SKILL.md",
            'for ((i=0;i>=0;i--)); do echo "$i"; done',
            'if [[ "$a" > "$b" ]]; then echo y; fi',
            'echo $(( n > 3 ))',
            "# don't write a > b.txt\ncat SKILL.md",
            "cat > /tmp/out.txt <<'EOF'\nif a > b.txt\nEOF",
        ]:
            with self.subTest(command=command[:60]):
                self.assertIsNone(self.verdict(command, self.V2))

    def test_variable_paths_resolve_before_judging(self):
        os.environ['HOOK_TEST_OUT'] = self.SP
        self.addCleanup(os.environ.pop, 'HOOK_TEST_OUT', None)
        for command in [f'SP={self.SP}; ls > "$SP/x.txt"', f'export SP={self.SP}; ls > ${{SP}}/x.txt',
                        'ls > "$HOOK_TEST_OUT/x.txt"', f'SP={self.SP}; S2="$SP/a"; ls > "$S2/x.txt"',
                        f'D={self.SP}; cp SKILL.md "$D/x.txt"']:
            with self.subTest(command=command):
                self.assertIsNone(self.verdict(command, self.V2))

    def test_real_writes_are_still_blocked(self):
        v2 = self.V2
        inside = [
            'echo hi > a.txt', 'echo hi >> a.txt', 'echo hi>a.txt', 'echo n>=30', 'ls > "$SP_NOT_SET_ANYWHERE/x.txt"',
            'ls > $(mktemp)', 'T=$(mktemp); ls > "$T"', 'X=a.txt; echo hi > "$X"', 'cat > a.txt <<EOF\nbody\nEOF',
            "python3 -c \"import os; os.system('echo hi > a.txt')\"", "bash -c 'echo hi > a.txt'", 'eval "echo hi > a.txt"',
            "awk '{print > \"out\"}' SKILL.md", 'x="$(echo hi > out)"', 'bash <<EOF\necho hi > out\nEOF',
            "cat <<EOF | sh\necho hi > out\nEOF", "sed -i '' s/a/b/ SKILL.md", "echo 'unterminated > a.txt",
        ]
        for command in inside:
            with self.subTest(command=command):
                self.assertIn('组件只读', self.verdict(command, v2) or '')
        anywhere = [
            f'echo hi > {v2}/a.txt', f"echo hi >> '{v2}/a.txt'", f'P={v2}/a.txt; echo hi > "$P"',
            f'ROOT={self.root}; echo hi > "$ROOT/{v2}/a.txt"', f'cd {v2}; D={self.SP}; X=a.txt; echo hi > $X',
            f"python3 -c \"import os; os.system('echo hi > {v2}/a.txt')\"", f"bash -c 'echo hi > {v2}/a.txt'",
            f"python3 - <<'EOF'\nimport os\nos.system('echo hi > {v2}/a.txt')\nEOF",
            f'D={v2}; cp /tmp/a "$D/a.txt"', f'D={v2}; rm "$D/a.txt"', f'D={v2}; cd "$D"; rm a.txt',
        ]
        for command in anywhere:
            with self.subTest(command=command):
                self.assertIn('组件只读', self.verdict(command) or '')
        self.assertIn('只能由用户本人修改', self.verdict(f'P={hook.SWITCH}; echo x > "$P"') or '')

    def test_stack_rule_reads_heredocs_and_variables_the_same_way(self):
        (self.root / '岗位交接').mkdir()
        gate = lambda command: hook.stack_violation(self.root, {
            'tool_name': 'Bash', 'tool_input': {'command': command}, 'cwd': str(self.root / '岗位交接')})
        for command in [f'SP={self.SP}; ls > "$SP/x.txt"', "python3 - <<'EOF'\nfor i in x:\n    if i>=0: pass\nEOF"]:
            with self.subTest(command=command):
                self.assertIsNone(gate(command))
        for command in ['ls > x.txt', 'X=x.txt; ls > "$X"', 'ls > "$SP_NOT_SET_ANYWHERE/x.txt"',
                        'cat <<EOF > x.txt\nbody\nEOF']:
            with self.subTest(command=command):
                self.assertIn('记录唯一门禁', gate(command) or '')


class RedirectFormsAndLinkTargetsTests(unittest.TestCase):
    V2 = '.agents/skills/job-application-assistant-v2'
    WRITE_OPS = ('1>', '2>', '3>', '&>', '>|', '>&', '1>>', '2>>', '&>>')

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name).resolve()
        (self.root / '产品管理').mkdir()

    def tearDown(self):
        self.tmp.cleanup()

    def verdict(self, command, cwd=None):
        return hook.component_violation(self.root, {
            'tool_name': 'Bash', 'tool_input': {'command': command}, 'session_id': 'locked',
            'cwd': str(self.root / cwd) if cwd else str(self.root)})

    def assert_blocked(self, commands, cwd=None):
        for command in commands:
            with self.subTest(command=command):
                self.assertIn('组件只读', self.verdict(command, cwd) or '')

    def assert_passes(self, commands, cwd=None):
        for command in commands:
            with self.subTest(command=command):
                self.assertIsNone(self.verdict(command, cwd))

    def test_every_redirect_form_writing_to_a_locked_path_is_blocked(self):
        v2 = self.V2
        self.assert_blocked([f'echo hi {op} {v2}/a.txt' for op in self.WRITE_OPS])
        self.assert_blocked([f'echo hi {op}{v2}/a.txt' for op in self.WRITE_OPS])
        self.assert_blocked([f'echo hi {op} a.txt' for op in self.WRITE_OPS], v2)
        self.assert_blocked([
            f'P={v2}/a.txt; echo hi 2> "$P"', f'P={v2}/a.txt; echo hi >| "$P"', f"echo hi 2> '{v2}/a.txt'",
            f'echo hi >| {v2}/a.txt | cat', f'ls 2>&1 >| {v2}/a.txt', f'ls /nope 2> {v2}/a.txt >/dev/null',
        ])
        self.assertIn('只能由用户本人修改', self.verdict(f'echo x 2> {hook.SWITCH}') or '')
        self.assertIn('只能由用户本人修改', self.verdict(f'echo x >| {hook.SWITCH}') or '')

    def test_redirects_inside_text_that_runs_are_blocked(self):
        v2 = self.V2
        self.assert_blocked([f"bash -c 'echo hi {op} {v2}/a.txt'" for op in self.WRITE_OPS])

    def test_descriptor_and_null_redirects_are_not_writes(self):
        self.assert_passes([
            'ls 2>&1', 'ls 2>&1 | head -3', 'ls >&2', 'ls 1>&2', 'ls 2>&-', 'ls >&-', 'ls >/dev/null 2>&1',
            'ls 2>/dev/null', 'ls 2> /dev/null', 'ls 1>/dev/null', 'ls &>/dev/null', 'ls &> /dev/null',
            'ls >|/dev/null', 'ls 2>>/dev/null',
            'echo hi 2> /tmp/x.txt', 'echo hi >| /tmp/x.txt', 'echo hi &> /tmp/x.txt', 'echo hi >& /tmp/x.txt',
            'cat SKILL.md | head -3', 'ls || echo none', 'mkdir -p /tmp/x 2>/dev/null', 'rm -f /tmp/x 2>/dev/null',
            "bash -c 'ls 2>&1'", "bash -c 'echo hi >&2'", 'bash -c "ls >/dev/null 2>&1"',
        ], self.V2)

    def test_redirect_characters_in_data_are_not_writes(self):
        self.assert_passes(["grep -n '2> x.txt' SKILL.md", "grep -n '>|' SKILL.md", 'echo $(( 2>1 ))',
                            'if [[ 2>1 ]]; then echo y; fi', "# echo 2> a.txt\ncat SKILL.md",
                            "cat > /tmp/out.txt <<'EOF'\necho 2> a.txt\necho >| a.txt\nEOF"], self.V2)

    def test_a_trailing_redirect_does_not_hide_the_real_target(self):
        v2 = self.V2
        self.assert_blocked([f'ln -s /tmp/x {v2}/a {tail}' for tail in ('2>/dev/null', '> /dev/null', '2>&1', '>/dev/null 2>&1')])
        self.assert_blocked([f'cp /tmp/x {v2}/a 2>/dev/null', f'cp /tmp/x {v2}/a > /dev/null'])

    def test_link_sources_are_not_writes(self):
        v2, abs_v2 = self.V2, f'{self.root}/{self.V2}'
        self.assert_passes([
            f'ln -s {v2}/SKILL.md /tmp/x', f'ln -sf {abs_v2}/SKILL.md /tmp/x', f'ln {v2}/SKILL.md /tmp/x',
            f'ln -s {v2}/a {v2}/b /tmp/dir', f'ln -s {v2}/SKILL.md', f'ln -s {v2}/SKILL.md /tmp/x 2>/dev/null',
            f'ln -s {v2}/SKILL.md -t /tmp', f'ln -t /tmp -s {v2}/SKILL.md', f'ln -st /tmp {v2}/SKILL.md',
            f'ln --target-directory=/tmp {v2}/SKILL.md', f'ln --target-directory /tmp {v2}/SKILL.md',
        ])

    def test_link_locations_are_still_writes(self):
        v2 = self.V2
        self.assert_blocked([
            f'ln -s /tmp/x {v2}/a', f'ln -sf {v2}/a {v2}/b', f'ln /tmp/x {v2}/a', f'ln -sfn /tmp/x {v2}',
            f'ln -s /tmp/a /tmp/b {v2}/', f'ln -s /tmp/x -t {v2}', f'ln -t {v2} /tmp/x', f'ln -st {v2} /tmp/x',
            f'ln -t{v2} /tmp/x', f'ln --target-directory={v2} /tmp/x', f'ln --target-directory {v2} /tmp/x',
        ])
        self.assert_blocked(['ln -s /tmp/x', 'ln -s /tmp/x a', 'ln -s /tmp/x -t .', f'cd {self.root}/{v2}; ln -s /tmp/x'], v2)
        self.assertIn('只能由用户本人修改', self.verdict(f'ln -s /tmp/x {hook.SWITCH}') or '')

    def stack_gate(self, command, cwd='岗位交接'):
        (self.root / cwd).mkdir(exist_ok=True)
        return hook.stack_violation(self.root, {
            'tool_name': 'Bash', 'tool_input': {'command': command}, 'cwd': str(self.root / cwd)})

    def test_stack_rule_sees_every_redirect_form(self):
        for command in [f'echo hi {op} {target}' for op in self.WRITE_OPS + ('>', '>>')
                        for target in ('x.txt', '"x.txt"')] + [
            'echo hi >|x.txt', 'echo hi &>x.txt', 'echo hi>&x.txt', "bash -c 'echo hi >| x.txt'",
            'y=$(echo hi &> x.txt)', 'X=x.txt; echo hi >| "$X"', 'ls 2>&1 >| x.txt',
            'cp /tmp/a x.txt 2>/dev/null', 'cp /tmp/a x.txt 2>&1', 'cp /tmp/a x.txt >/dev/null 2>&1',
        ]:
            with self.subTest(command=command):
                self.assertIn('记录唯一门禁', self.stack_gate(command) or '')
        self.assertIn('记录唯一门禁', self.stack_gate('echo hi &> x-副本.txt', '产品管理') or '')

    def test_stack_rule_ignores_descriptors_null_and_quoted_operators(self):
        for command in [
            'ls 2>&1', 'ls 2>&1 | head -3', 'ls >&2', 'ls 1>&2', 'ls 2>&-', 'ls >&-', 'ls >/dev/null 2>&1',
            'ls 2>/dev/null', 'ls 1>/dev/null', 'ls &>/dev/null', 'ls &>>/dev/null', 'ls >|/dev/null', 'ls >&/dev/null',
            'echo hi &> /tmp/x.txt', 'echo hi >| /tmp/x.txt', 'echo hi >& /tmp/x.txt',
            "grep -n '>|' x.txt", 'grep -n "&>" x.txt', "echo '>&' x.txt", "echo '>' x.txt", "printf '%s' \">>\" x.txt",
            "cat <<'EOF'\necho hi &> out\necho hi >| out\necho hi >& out\nEOF",
        ]:
            with self.subTest(command=command):
                self.assertIsNone(self.stack_gate(command))

    def test_writing_through_an_outside_link_is_a_known_gap(self):
        # Targets are normalized, not resolved, so a link made outside the lock
        # can be written through. Tightening this needs symlink resolution.
        self.assertIsNone(self.verdict(f'ln -s {self.V2} /tmp/x; echo hi > /tmp/x/a.txt'))
