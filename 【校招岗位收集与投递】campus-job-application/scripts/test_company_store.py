import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import workbook_store  # noqa: E402
import company_store as C  # noqa: E402

SCRIPT = HERE / 'company_store.py'
HEAD = ['地区', '公司', '主要行业', '企业性质', '岗位名称', '工作地点', '岗位JD', '所属BU', '岗位链接', '投递状态', '备注']


class CompanyStoreTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        r = self.root = Path(self.tmp.name)
        (r / '产品管理').mkdir()
        (r / '岗位交接').mkdir()
        wb = Workbook()
        ws = wb.active
        ws.title = '中国大陆'
        ws.append(HEAD)
        ws.append(['上海', '甲公司', '互联网', '私企', None, None, None, None, None, '已有岗位', None])
        ws.append(['上海', '甲公司', '互联网', '私企', 'AI产品经理', '上海', 'JD', '', 'https://jia.jobs.example.cn/campus/position/123456/detail', '岗位待投递', None])
        ws.append(['上海', '甲公司', '互联网', '私企', '游戏产品', '上海', 'JD', '', 'https://game.jia.example.cn/job/778899', '岗位待投递', None])
        ws.append(['北京', '乙公司', '互联网', '私企', None, None, None, None, None, '已有岗位', None])
        ws.append(['北京', '乙公司', '互联网', '私企', '产品经理', '北京', 'JD', '', 'https://www.nowcoder.com/jobs/detail/5566', '岗位待投递', None])
        ws.append(['北京', '丙公司', '互联网', '私企', None, None, None, None, None, '已有岗位', None])
        ws.append(['北京', '丙公司', '互联网', '私企', 'A岗', '北京', 'JD', '', 'https://bing.example.cn', '岗位待投递', None])
        ws.append(['北京', '丙公司', '互联网', '私企', 'B岗', '北京', 'JD', '', 'https://bing.example.cn', '岗位待投递', None])
        ws.append(['北京', '丁公司', '互联网', '私企', None, None, None, None, None, '待搜索', None])
        wb.save(r / 'book.xlsx')
        (r / '产品管理/company-labels.json').write_text(json.dumps({"schemaVersion": 1, "rowKind": "company",
            "labels": [{"value": "待搜索"}, {"value": "已有岗位"}]}, ensure_ascii=False), encoding='utf-8')
        cfg = {"schemaVersion": 1, "rowKind": "role", "labelGroups": ["产品管理/company-labels.json"], "column": "投递状态", "sheets": ["中国大陆"],
               "labels": [{"value": "岗位待投递"}, {"value": "岗位已投递"}, {"value": "岗位失效/关闭"}],
               "confirmationLedger": "产品管理/确认.json", "masterData": "岗位交接/工作簿全量.json", "workbook": "book.xlsx",
               "companyStore": {"dir": "岗位交接/公司", "batchFile": "岗位交接/投递批次.json", "forwardReview": "岗位交接/正向评审.html",
                                "thirdPartyHosts": ["nowcoder.com"]}}
        (r / '产品管理/labels.json').write_text(json.dumps(cfg, ensure_ascii=False), encoding='utf-8')
        (r / '岗位交接/工作簿全量.json').write_text(json.dumps(workbook_store.snapshot(r / 'book.xlsx'), ensure_ascii=False), encoding='utf-8')
        # 旧版公司文件：带研究记录、方案原文、状态副本和内部编号，refresh 后必须全部消失
        old = r / '岗位交接/公司/中国大陆/甲公司'
        old.mkdir(parents=True)
        (old / '公司.json').write_text(json.dumps({"companyId": "jia", "公司": "甲公司", "地区": "中国大陆", "origins": [],
            "研究记录": [{"原文": "旧结论"}], "搜索方案原文": [{"批次": "x"}], "投递窗口": "旧窗口",
            "roles": [{"主表行身份": {"sheet": "中国大陆", "company": "甲公司", "link": "https://jia.jobs.example.cn/campus/position/123456/detail", "title": "AI产品经理"},
                       "岗位ID": "r-jia-ai", "投递状态": "旧状态", "顺序": 1, "用户决定": "apply"}]}, ensure_ascii=False), encoding='utf-8')
        stale = r / '岗位交接/公司/中国大陆/只有备注的公司'
        stale.mkdir(parents=True)
        (stale / '公司.json').write_text('{"companyId":"x","公司":"只有备注的公司","地区":"中国大陆","roles":[]}', encoding='utf-8')

    def tearDown(self):
        self.tmp.cleanup()

    def run_cmd(self, *args, ok=True):
        p = subprocess.run([sys.executable, '-B', str(SCRIPT), '--workspace', str(self.root), '--labels', '产品管理/labels.json', *args],
                           capture_output=True, text=True, encoding='utf-8',
                           env={**os.environ, 'CODEX_THREAD_ID': 'test-session', 'CODEX_TURN_ID': 'test-turn', 'DSH_SESSION_ID': ''})
        if ok and p.returncode:
            self.fail(p.stderr + p.stdout)
        return json.loads(p.stdout) if p.stdout.strip() and ok else p

    def test_role_id_only_from_link(self):
        self.assertEqual(C.Store.role_id('https://talent.baidu.com/jobs/detail/GRADUATE/0aa80bfa-a052-4052-afc1-28c38407c13c'), '0aa80bfa-a052-4052-afc1-28c38407c13c')
        self.assertEqual(C.Store.role_id('https://x.cn/pb/posDetail.html?postId=6a83c54768cc&postType=campus'), '6a83c54768cc')
        self.assertIsNone(C.Store.role_id('https://sensetime.jobs.feishu.cn'))
        self.assertIsNone(C.Store.role_id('https://careers.pddglobalhr.com/campus/grad/detail?jobId=temu-pm'))
        self.assertEqual(C.Store.role_id('https://hr.4399om.com/main/?r=UserCenter/addValidate&jobID=JO20260728001'), 'JO20260728001')
        self.assertEqual(C.Store.role_id('https://maker.haier.net/client/campus/deliverfirst/id/60/fid/23/rid/392.html'), '392')
        self.assertEqual(C.Store.role_id('https://www.lixiang.com/employ/detail/19719.html?jobCode=A98366'), 'A98366')

    def test_explicit_requirements_keep_both_roles_and_clear_only_after_batch(self):
        import application_requirements as R
        self.run_cmd('refresh')
        self.assertEqual(self.set_order(('游戏产品', 1), ('AI产品经理', 2)).returncode, 0)
        req = R.Requirements(self.root, 'test-session', 'test-turn')
        req.set({'公司': ['甲公司'], '浏览器': 'iab'}, '只做甲公司，两个岗位都要，只用内置浏览器')
        original = (self.root / '岗位交接/工作簿全量.json').read_bytes()
        out = self.run_cmd('forward')
        self.assertEqual((out['companies'], out['roles'], out['excluded']), (1, 2, 0))
        self.assertTrue(out['requirementsCleared'])
        batch = json.loads((self.root / '岗位交接/投递批次.json').read_text(encoding='utf-8'))
        self.assertEqual([c['公司'] for c in batch['companies']], ['甲公司'])
        self.assertEqual(batch['execution']['browserChannels'], ['iab'])
        self.assertEqual([r['顺序'] for r in batch['companies'][0]['岗位']], [1, 2])
        self.assertEqual(json.loads(req.temp.read_text(encoding='utf-8')), {})
        self.assertEqual((self.root / '岗位交接/工作簿全量.json').read_bytes(), original)

    def test_invalid_explicit_company_does_not_consume_or_replace_batch(self):
        import application_requirements as R
        self.run_cmd('refresh')
        self.run_cmd('forward')
        batch = self.root / '岗位交接/投递批次.json'
        before = batch.read_bytes()
        req = R.Requirements(self.root, 'test-session', 'test-turn')
        req.set({'公司': ['不存在的公司']}, '只做指定公司')
        result = self.run_cmd('forward', ok=False)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn('指定公司没有待投递岗位', result.stderr)
        self.assertEqual(batch.read_bytes(), before)
        self.assertEqual(json.loads(req.temp.read_text(encoding='utf-8'))['requirements']['公司'], ['不存在的公司'])

    def test_refresh_overwrites_and_keeps_only_runtime_fields(self):
        out = self.run_cmd('refresh')
        self.assertEqual(out['folders_removed'], 1)
        self.assertFalse((self.root / '岗位交接/公司/中国大陆/只有备注的公司').exists())
        self.assertFalse((self.root / '岗位交接/公司/中国大陆/丁公司').exists())
        jia = json.loads((self.root / '岗位交接/公司/中国大陆/甲公司/公司.json').read_text(encoding='utf-8'))
        self.assertEqual(set(jia), {'schemaVersion', 'companyId', '公司', '地区', 'origins', '志愿分组依据', 'roles', 'updatedAt'})
        self.assertEqual(jia['companyId'], 'jia')
        self.assertEqual(jia['origins'], ['https://game.jia.example.cn', 'https://jia.jobs.example.cn'])
        self.assertEqual(jia['roles'][0], {'主表行身份': jia['roles'][0]['主表行身份'], '岗位ID': '123456', 'BU': '未披露', '志愿分组': '公司', '顺序': 1})
        self.assertNotIn('旧', json.dumps(jia, ensure_ascii=False))

    def set_order(self, *pairs):
        f = self.root / 'order.json'
        f.write_text(json.dumps([{"sheet": "中国大陆", "company": "甲公司", "title": t, "顺序": n} for t, n in pairs], ensure_ascii=False), encoding='utf-8')
        return self.run_cmd('set-order', '--input', str(f), '--quote', '游戏产品第一，AI第二', ok=False)

    def sync(self, *args, ok=True):
        p = subprocess.run([sys.executable, '-B', str(HERE / 'json_excel_sync.py'), '--workspace', str(self.root), '--labels', '产品管理/labels.json', *args],
                           capture_output=True, text=True, encoding='utf-8')
        if ok and p.returncode:
            self.fail(p.stderr + p.stdout)
        return p

    def test_relink_keeps_user_order_and_is_blocked_after_handoff(self):
        """换链接后 refresh 保留用户给的志愿顺序；岗位在待评审批次里提示须重新 forward；批次已交接后拒绝换链接。"""
        self.run_cmd('refresh')
        self.assertEqual(self.set_order(('游戏产品', 1), ('AI产品经理', 2)).returncode, 0)
        f = self.root / 'relink.json'
        old, new, newer = ('https://game.jia.example.cn/job/778899', 'https://game.jia.example.cn/job/778900', 'https://game.jia.example.cn/job/778901')
        f.write_text(json.dumps([{"sheet": "中国大陆", "company": "甲公司", "link": old, "title": "游戏产品", "new_link": new}], ensure_ascii=False), encoding='utf-8')
        self.sync('relink', '--input', str(f), '--quote', '游戏产品换官网新链接')
        self.run_cmd('refresh')
        jia = json.loads((self.root / '岗位交接/公司/中国大陆/甲公司/公司.json').read_text(encoding='utf-8'))
        self.assertEqual([(r['岗位ID'], r.get('顺序')) for r in jia['roles']], [('123456', 2), ('778900', 1)])
        self.assertEqual(self.run_cmd('forward')['roles'], 2)
        f.write_text(json.dumps([{"sheet": "中国大陆", "company": "甲公司", "link": new, "title": "游戏产品", "new_link": newer}], ensure_ascii=False), encoding='utf-8')
        self.assertIn('须重新生成批次', self.sync('relink', '--input', str(f), '--preview').stdout)
        self.run_cmd('forward-confirm', '--quote', '确认交接')
        self.assertIn('已交接', self.sync('relink', '--input', str(f), '--quote', 'x', ok=False).stderr)

    def test_link_exception_lets_list_page_link_through_forward(self):
        """列表页链接默认不交接；用户确认链接例外后交接，并带上定位说明。"""
        self.run_cmd('refresh')
        self.assertEqual(self.set_order(('游戏产品', 1), ('AI产品经理', 2)).returncode, 0)
        old, lst = 'https://jia.jobs.example.cn/campus/position/123456/detail', 'https://jia.jobs.example.cn/campus-recruit-list.html?positionId=123456'
        f = self.root / 'in.json'
        f.write_text(json.dumps([{"sheet": "中国大陆", "company": "甲公司", "link": old, "title": "AI产品经理", "new_link": lst}], ensure_ascii=False), encoding='utf-8')
        self.sync('relink', '--input', str(f), '--quote', '换成官网列表链接')
        self.run_cmd('refresh')
        self.run_cmd('forward')
        b = json.loads((self.root / '岗位交接/投递批次.json').read_text(encoding='utf-8'))
        self.assertEqual({(e['公司'], e.get('岗位')): e['原因'] for e in b['excluded']}[('甲公司', 'AI产品经理')], C.LIST_PAGE_PROBLEM)
        f.write_text(json.dumps([{"sheet": "中国大陆", "company": "甲公司", "link": lst, "title": "AI产品经理",
                                  "定位说明": "列表页按 positionId 定位，点岗位名展开 JD"}], ensure_ascii=False), encoding='utf-8')
        self.sync('link-exception', '--input', str(f), '--quote', '官网没有详情页，这个可以')
        self.run_cmd('forward')
        b = json.loads((self.root / '岗位交接/投递批次.json').read_text(encoding='utf-8'))
        jia = next(c for c in b['companies'] if c['公司'] == '甲公司')
        self.assertEqual([(r['岗位名称'], r['岗位ID'], r.get('链接说明')) for r in jia['岗位']],
                         [('游戏产品', '778899', None), ('AI产品经理', '123456', '列表页按 positionId 定位，点岗位名展开 JD')])
        self.assertIn('链接例外：列表页按 positionId 定位', (self.root / '岗位交接/正向评审.html').read_text(encoding='utf-8'))

    def test_forward_requires_user_given_order_within_each_group(self):
        self.run_cmd('refresh')
        out = self.run_cmd('forward')
        self.assertEqual(out['roles'], 0)
        b = json.loads((self.root / '岗位交接/投递批次.json').read_text(encoding='utf-8'))
        self.assertIn('志愿顺序未由用户给全', {(e['公司'], e.get('岗位')): e['原因'] for e in b['excluded']}[('甲公司', 'AI产品经理')])
        self.assertNotEqual(self.set_order(('游戏产品', 1), ('AI产品经理', 1)).returncode, 0)
        self.assertNotEqual(self.set_order(('游戏产品', 0), ('AI产品经理', 3)).returncode, 0)
        self.assertEqual(self.set_order(('游戏产品', 1), ('AI产品经理', 2)).returncode, 0)
        jia = json.loads((self.root / '岗位交接/公司/中国大陆/甲公司/公司.json').read_text(encoding='utf-8'))
        self.assertEqual(jia['志愿顺序确认']['userQuote'], '游戏产品第一，AI第二')
        self.run_cmd('refresh')  # 重建保留用户给定的顺序
        b = self.run_cmd('forward') and json.loads((self.root / '岗位交接/投递批次.json').read_text(encoding='utf-8'))
        self.assertEqual([(r['岗位名称'], r['顺序'], r['志愿分组']) for r in b['companies'][0]['岗位']],
                         [('游戏产品', 1, '公司'), ('AI产品经理', 2, '公司')])

    def test_application_url_written_back_by_v2_survives_refresh_and_reaches_batch(self):
        self.run_cmd('refresh')
        f = self.root / '岗位交接/公司/中国大陆/甲公司/公司.json'
        doc = json.loads(f.read_text(encoding='utf-8'))
        doc['roles'][0]['投递链接'] = 'https://jia.jobs.example.cn/campus/apply/123456'
        f.write_text(json.dumps(doc, ensure_ascii=False), encoding='utf-8')
        self.run_cmd('refresh')
        self.assertEqual(json.loads(f.read_text(encoding='utf-8'))['roles'][0]['投递链接'], 'https://jia.jobs.example.cn/campus/apply/123456')

    def test_bu_parallel_quota_splits_order_groups(self):
        m = json.loads((self.root / '岗位交接/工作簿全量.json').read_text(encoding='utf-8'))
        sh = m['sheets'][0]
        for r in sh['rows']:
            if r['row'] in (3, 4):  # 第2行是甲公司的公司行，岗位从第3行开始
                r['cells'].append({'column': 8, 'value': '甲BU' if r['row'] == 3 else '乙BU', 'type': 's'})
                r['cells'].sort(key=lambda c: c['column'])
        head = ['公司', '范围', '适用计划/批次', '投递上限', '当前投递数量', '是否分BU可平行投递', '是否可以滚动投递', '规则原文', '来源/核验日期']
        vals = ['甲公司', '公司总量', '全部计划', 2, 0, '是', '否', '原文', '来源：甲公司-FAQ 核验日期：2026-10-20']
        m['sheets'].append({'name': '投递名额', 'rows': [{'row': 1, 'cells': [{'column': i + 1, 'value': h, 'type': 's'} for i, h in enumerate(head)]},
                            {'row': 2, 'cells': [{'column': i + 1, 'value': v, 'type': 's' if isinstance(v, str) else 'n'} for i, v in enumerate(vals)]}],
                            'max_row': 2, 'max_column': len(head), 'state': 'visible', 'merged_ranges': [], 'tables': []})
        (self.root / '岗位交接/工作簿全量.json').write_text(json.dumps(m, ensure_ascii=False), encoding='utf-8')
        self.run_cmd('refresh')
        jia = json.loads((self.root / '岗位交接/公司/中国大陆/甲公司/公司.json').read_text(encoding='utf-8'))
        self.assertEqual([r['志愿分组'] for r in jia['roles']], ['甲BU', '乙BU'])
        self.assertEqual(self.run_cmd('forward')['roles'], 2)  # 每个分组只有一个岗位，不需要顺序

    def test_forward_rejects_non_role_links_and_has_no_history(self):
        self.run_cmd('refresh')
        self.set_order(('AI产品经理', 1), ('游戏产品', 2))
        out = self.run_cmd('forward')
        self.assertEqual(out['roles'], 2)
        b = json.loads((self.root / '岗位交接/投递批次.json').read_text(encoding='utf-8'))
        self.assertNotIn('history', b)
        reasons = {(e['公司'], e.get('岗位')): e['原因'] for e in b['excluded']}
        self.assertIn('第三方', reasons[('乙公司', '产品经理')])
        self.assertIn('共用同一链接', reasons[('丙公司', 'A岗')])
        self.assertIn('没有官方岗位ID', reasons[('丙公司', 'B岗')])
        self.assertEqual(self.run_cmd('batch-status')['handoffProblem'], '投递批次未经用户正向评审确认')
        self.run_cmd('forward-confirm', '--quote', '确认交接')
        self.assertIsNone(self.run_cmd('batch-status')['handoffProblem'])
        b = json.loads((self.root / '岗位交接/投递批次.json').read_text(encoding='utf-8'))
        b['companies'][0]['岗位'].append({'岗位ID': 'x'})
        (self.root / '岗位交接/投递批次.json').write_text(json.dumps(b, ensure_ascii=False), encoding='utf-8')
        self.assertIn('被改动', self.run_cmd('batch-status')['handoffProblem'])
        self.assertNotEqual(self.run_cmd('forward', ok=False).returncode, 0)


    def outcome(self, key, event, title, url, rid):
        f = self.root / '岗位交接/投递回传.json'
        d = json.loads(f.read_text(encoding='utf-8')) if f.exists() else {'schemaVersion': 1, 'roles': {}}
        d['roles'][key] = {'event': event, 'companyId': key.split('|')[0], 'roleId': rid, 'roleUrl': url, 'roleTitle': title, 'batchId': 'B'}
        f.write_text(json.dumps(d, ensure_ascii=False), encoding='utf-8')

    def test_reverse_review_writes_master_only_after_user_confirmation(self):
        self.run_cmd('refresh')
        self.set_order(('AI产品经理', 1), ('游戏产品', 2))
        self.run_cmd('forward')
        self.assertNotEqual(self.run_cmd('reverse', ok=False).returncode, 0)  # 未确认交接不能反向评审
        self.run_cmd('forward-confirm', '--quote', '确认交接')
        self.outcome('jia|123456', 'submission_confirmed', 'AI产品经理', 'https://jia.jobs.example.cn/campus/position/123456/detail', '123456')
        self.outcome('jia|778899', 'submission_uncertain', '游戏产品', 'https://game.jia.example.cn/job/778899', '778899')
        out = self.run_cmd('reverse')
        self.assertEqual(out['show_user_exactly'], ['中国大陆｜甲公司｜AI产品经理：当前状态是「岗位待投递」，更新状态是「岗位已投递」'])
        self.assertIn('游戏产品：投递结果不确定或受阻，处理前禁止重跑', out['notes'])
        master = (self.root / '岗位交接/工作簿全量.json').read_text(encoding='utf-8')
        self.assertNotIn('岗位已投递', master)  # reverse 只生成评审，不改主表
        self.assertNotEqual(self.run_cmd('forward', ok=False).returncode, 0)  # 未回写完不能开下一批
        res = self.run_cmd('reverse-confirm', '--quote', '确认回写')
        self.assertEqual((res['state'], res['statusChanges'], res['unresolved']), ('等待反向评审', 1, ['jia|778899']))
        self.assertIn('岗位已投递', (self.root / '岗位交接/工作簿全量.json').read_text(encoding='utf-8'))
        self.run_cmd('reverse')
        res = self.run_cmd('reverse-confirm', '--quote', '不确定那个我已核实没投上', '--resolved', 'jia|778899')
        self.assertEqual(res['state'], '已回写')
        self.assertFalse((self.root / '岗位交接/投递回传.json').exists())


if __name__ == '__main__':
    unittest.main()
