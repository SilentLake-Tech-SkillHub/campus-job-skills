from test_browser_capabilities import fixture_environment
from unittest.mock import patch
import os
import json
from pathlib import Path
import tempfile
import unittest

import application_requirements as R
import json_excel_sync as S


class RequirementsTests(unittest.TestCase):
    def test_stale_legacy_lock_is_recovered_and_live_owner_is_preserved(self):
        import os, time
        lock = self.root / '.runtime/user_application_requirements.lock'
        lock.parent.mkdir(parents=True, exist_ok=True)
        lock.write_text('')
        os.utime(lock, (time.time() - 1200, time.time() - 1200))
        with self.store.locked():
            self.assertEqual(json.loads(lock.read_text())['pid'], os.getpid())
        self.assertFalse(lock.exists())
        lock.write_text(json.dumps({'pid': os.getpid(), 'createdAt': time.time()}))
        with self.assertRaisesRegex(ValueError, '另一操作'):
            with self.store.locked():
                self.fail('active lock must never be acquired')
        self.assertTrue(lock.exists())

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        environment=patch.dict(os.environ,fixture_environment(self.root));environment.start();self.addCleanup(environment.stop)
        self.store = R.Requirements(self.root, 'session-a', 'turn-1')

    def test_same_turn_default_does_not_ask_again_next_turn_resets(self):
        self.assertTrue(self.store.begin()['askUserRequirements'])
        self.store.set({}, '默认')
        self.assertFalse(self.store.begin()['askUserRequirements'])
        self.assertFalse(self.store.begin()['askUserRequirements'])
        next_turn = R.Requirements(self.root, 'session-a', 'turn-2')
        self.assertTrue(next_turn.begin()['askUserRequirements'])

    def test_overwrite_and_consume_only_after_success(self):
        self.store.set({'公司': ['甲公司']}, '只做甲公司')
        self.store.set({'公司': ['乙公司'], '浏览器': 'iab'}, '改为乙公司，只用内置浏览器')
        seen = []
        def build(req):
            seen.append(req)
            return {'batchId': 'batch-a'}
        self.store.build(build)
        self.assertEqual(seen, [{'公司': ['乙公司'], '浏览器': 'iab'}])
        self.assertEqual(R.read(self.store.temp), {})
        self.store.build(build)
        self.assertEqual(seen[-1], {})

    def test_failed_build_keeps_requirements_for_retry(self):
        self.store.set({'公司': ['甲公司']}, '甲公司')
        def fail(_):
            raise ValueError('review write failed')
        with self.assertRaisesRegex(ValueError, 'review write failed'):
            self.store.build(fail)
        self.assertEqual(R.read(self.store.temp)['requirements'], {'公司': ['甲公司']})

    def test_other_conversation_cannot_read_overwrite_or_consume_request(self):
        self.store.set({'公司': ['甲公司']}, '甲公司')
        other = R.Requirements(self.root, 'session-b', 'turn-1')
        for operation in [other.begin, lambda: other.set({}, '默认'), lambda: other.build(lambda _: {'batchId': 'b'})]:
            with self.assertRaisesRegex(ValueError, '另一个对话'):
                operation()

    def test_previous_turn_parameters_not_reused(self):
        self.store.set({'公司': ['甲公司']}, '甲公司')
        later = R.Requirements(self.root, 'session-a', 'turn-2')
        seen = []
        later.build(lambda req: (seen.append(req), {'batchId': 'b'})[1])
        self.assertEqual(seen, [{}])
        self.assertEqual(R.read(later.temp), {})

    def test_complete_preserves_next_cycle_request(self):
        self.store.set({'公司': ['甲公司']}, '甲公司')
        self.store.build(lambda _: {'batchId': 'batch-a'})
        self.store.set({'公司': ['乙公司']}, '下一批乙公司')
        self.assertTrue(self.store.complete('batch-a')['newerRequirementsPreserved'])
        self.assertEqual(R.read(self.store.temp)['requirements'], {'公司': ['乙公司']})

    def test_empty_table_is_default_and_completion_check_is_empty(self):
        got = self.store.build(lambda req: {'batchId': 'batch-a', 'parameters': req})
        self.assertEqual(got['parameters'], {})
        self.assertFalse(got['requirementsConsumed'])
        self.assertTrue(self.store.complete('batch-a')['requirementsEmpty'])

    def test_hook_resets_semiautomatic_round_and_is_idempotent_within_turn(self):
        self.store.begin()
        self.store.set({}, '默认')
        self.assertEqual(R.prompt_hook(self.root, {'hook_event_name': 'UserPromptSubmit', 'session_id': 'session-a', 'turn_id': 'turn-1'}), {})
        payload = {'hook_event_name': 'UserPromptSubmit', 'session_id': 'session-a', 'turn_id': 'turn-2'}
        self.assertIn('新用户回合', R.prompt_hook(self.root, payload)['hookSpecificOutput']['additionalContext'])
        self.assertEqual(R.prompt_hook(self.root, payload), {})
        self.assertTrue(R.Requirements(self.root, 'session-a').begin()['askUserRequirements'])

    def test_invalid_or_corrupt_request_never_silently_becomes_default(self):
        with self.assertRaisesRegex(ValueError, '重复'):
            self.store.set({'公司': ['甲公司', '甲公司']}, '甲公司')
        with self.assertRaisesRegex(ValueError, '不支持'):
            self.store.set({'投递状态': 'x'}, 'x')
        self.store.temp.parent.mkdir(parents=True)
        self.store.temp.write_text('{broken', encoding='utf-8')
        with self.assertRaisesRegex(ValueError, '停止'):
            self.store.begin()


    def test_maintenance_initialization_never_clears_a_pending_request(self):
        self.assertTrue(self.store.set({}, 'cleanup request', maintenance_clear=True)['maintenanceCleanup'])
        self.assertFalse(self.store.state.exists())
        self.store.set({'公司':['甲公司']}, 'new request')
        with self.assertRaisesRegex(ValueError, '不能清除'):
            self.store.set({}, 'cleanup request', maintenance_clear=True)


class SelectionTests(unittest.TestCase):
    """公司以公司编码为身份；需求里的公司名按公司主数据名称/别名完全一致对应到编码。"""
    def rows(self, n=12):
        rows, ledger, names = {}, {'entries': []}, {}
        for i in range(n):
            code = 'C%05d' % (i + 1)
            key = {'sheet': '合成地区', 'code': 'J%06d' % (i + 1)}
            r = {'kind': 'role', 'sheet': '合成地区', 'key': key, 'values': {'投递状态': '岗位待投递', '数据刷新时间': '2001-01-01'}}
            rows[code] = [r]
            names[code] = ['公司' + str(i)]
            ledger['entries'].append({'key': key, 'to': '岗位待投递', 'confirmed_at': '2002-01-%02dT00:00:00Z' % (i+1)})
        return rows, ledger, names

    def test_empty_requirements_select_latest_ten_companies(self):
        rows, ledger, names = self.rows()
        selected = R.select_companies(rows, ledger, {}, names)
        self.assertEqual(selected, ['C%05d' % (i + 1) for i in range(11, 1, -1)])

    def test_explicit_companies_override_default_limit_and_keep_requested_order(self):
        rows, ledger, names = self.rows(20)
        wanted = ['公司' + str(i) for i in range(20)]
        selected = R.select_companies(rows, ledger, {'公司': wanted}, names)
        self.assertEqual([names[c][0] for c in selected], wanted)

    def test_company_selection_does_not_slice_its_roles(self):
        rows, ledger, names = self.rows(1)
        rows['C00001'].append({'kind': 'role', 'sheet': '合成地区', 'key': {'sheet': '合成地区', 'code': 'J000099'},
                               'values': {'投递状态': '岗位待投递'}})
        selected = R.select_companies(rows, ledger, {'公司': ['公司0']}, names)
        self.assertEqual(len(rows[selected[0]]), 2)

    def test_missing_or_ambiguous_company_name_is_rejected(self):
        rows, ledger, names = self.rows(2)
        with self.assertRaisesRegex(ValueError, '没有待投递'):
            R.select_companies(rows, ledger, {'公司': ['不存在']}, names)
        names['C00002'].append('公司0')  # 主数据异常：一个名字对上两个编码
        with self.assertRaisesRegex(ValueError, '对上多个公司编码'):
            R.select_companies(rows, ledger, {'公司': ['公司0']}, names)
        self.assertEqual(R.select_companies(rows, ledger, {'公司': ['C00001']}, names), ['C00001'])

    def test_region_filters_roles(self):
        rows, ledger, names = self.rows(2)
        rows['C00002'][0]['sheet'] = '另一地区'
        self.assertEqual(R.select_companies(rows, ledger, {'地区': '另一地区'}, names), ['C00002'])

    def test_no_confirmation_uses_refresh_date_with_stable_ties(self):
        rows, _, names = self.rows(3)
        rows['C00002'][0]['values']['数据刷新时间'] = '2003-01-01'
        selected = R.select_companies(rows, {'entries': []}, {}, names)
        self.assertEqual(selected, ['C00002', 'C00001', 'C00003'])


if __name__ == '__main__':
    unittest.main()
