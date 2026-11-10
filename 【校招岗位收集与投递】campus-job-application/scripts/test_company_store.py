from test_browser_capabilities import fixture_environment
from unittest.mock import patch
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
import json_excel_sync as S  # noqa: E402
import master_codes as M  # noqa: E402

SCRIPT = HERE / 'company_store.py'
HEAD = ['地区', '公司', '主要行业', '企业性质', '岗位名称', '工作地点', '岗位JD', '所属BU', '岗位链接', '投递状态', '备注', '内推码']
AI, GAME = 'https://jia.jobs.example.cn/campus/position/123456/detail', 'https://game.jia.example.cn/job/778899'
# 编码按第一次出现的顺序分配：甲 C00001、乙 C00002、丙 C00003、丁 C00004；
# AI产品经理 J000001（站点 S0001 jia.jobs）、游戏产品 J000002（S0002 game.jia）、乙产品经理 J000003、丙A岗 J000004、丙B岗 J000005。
J_AI, J_GAME = 'J000001', 'J000002'


class CompanyStoreTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        r = self.root = Path(self.tmp.name)
        environment=patch.dict(os.environ,fixture_environment(self.root));environment.start();self.addCleanup(environment.stop)
        (r / '产品管理').mkdir()
        (r / '岗位交接').mkdir()
        wb = Workbook()
        ws = wb.active
        ws.title = '中国大陆'
        ws.append(HEAD)
        ws.append(['上海', '甲公司', '互联网', '私企', None, None, None, None, None, '已有岗位', None, None])
        ws.append(['上海', '甲公司', '互联网', '私企', 'AI产品经理', '上海', 'JD', '', AI, '岗位待投递', None, None])
        ws.append(['上海', '甲公司', '互联网', '私企', '游戏产品', '上海', 'JD', '', GAME, '岗位待投递', None, None])
        ws.append(['北京', '乙公司', '互联网', '私企', None, None, None, None, None, '已有岗位', None, None])
        ws.append(['北京', '乙公司', '互联网', '私企', '产品经理', '北京', 'JD', '', 'https://www.nowcoder.com/jobs/detail/5566', '岗位待投递', None, None])
        ws.append(['北京', '丙公司', '互联网', '私企', None, None, None, None, None, '已有岗位', None, None])
        ws.append(['北京', '丙公司', '互联网', '私企', 'A岗', '北京', 'JD', '', 'https://bing.example.cn', '岗位待投递', None, None])
        ws.append(['北京', '丙公司', '互联网', '私企', 'B岗', '北京', 'JD', '', 'https://bing.example.cn', '岗位待投递', None, None])
        ws.append(['北京', '丁公司', '互联网', '私企', None, None, None, None, None, '待搜索', None, None])
        wb.save(r / 'book.xlsx')
        (r / '产品管理/company-labels.json').write_text(json.dumps({"schemaVersion": 1, "rowKind": "company",
            "labels": [{"value": "待搜索"}, {"value": "已有岗位"}]}, ensure_ascii=False), encoding='utf-8')
        cfg = {"schemaVersion": 1, "rowKind": "role", "labelGroups": ["产品管理/company-labels.json"], "column": "投递状态", "sheets": ["中国大陆"],
               "labels": [{"value": "岗位待投递"}, {"value": "岗位已投递"}, {"value": "岗位失效/关闭"}],
               "confirmationLedger": "产品管理/确认.json", "masterData": "岗位交接/工作簿全量.json", "workbook": "book.xlsx",
               "companyStore": {"dir": "岗位交接/公司", "forwardReview": "岗位交接/正向评审.html",
                                "thirdPartyHosts": ["nowcoder.com"]}}
        (r / '产品管理/labels.json').write_text(json.dumps(cfg, ensure_ascii=False), encoding='utf-8')
        snap = workbook_store.snapshot(r / 'book.xlsx')
        M.assign_missing_codes(snap, ['中国大陆'])
        S.refresh_digest(snap)
        (r / '岗位交接/工作簿全量.json').write_text(json.dumps(snap, ensure_ascii=False), encoding='utf-8')

    def tearDown(self):
        self.tmp.cleanup()

    def run_cmd(self, *args, ok=True):
        p = subprocess.run([sys.executable, '-B', str(SCRIPT), '--workspace', str(self.root), '--labels', '产品管理/labels.json', *args],
                           capture_output=True, text=True, encoding='utf-8',
                           env={**os.environ, 'CODEX_THREAD_ID': 'test-session', 'CODEX_TURN_ID': 'test-turn', 'DSH_SESSION_ID': ''})
        if ok and p.returncode:
            self.fail(p.stderr + p.stdout)
        return json.loads(p.stdout) if p.stdout.strip() and ok else p

    def sync(self, *args, ok=True):
        p = subprocess.run([sys.executable, '-B', str(HERE / 'json_excel_sync.py'), '--workspace', str(self.root), '--labels', '产品管理/labels.json', *args],
                           capture_output=True, text=True, encoding='utf-8')
        if ok and p.returncode:
            self.fail(p.stderr + p.stdout)
        return p

    def write(self, name, data):
        f = self.root / name
        f.write_text(json.dumps(data, ensure_ascii=False), encoding='utf-8')
        return str(f)

    def doc(self, code='C00001'):
        return json.loads((self.root / f'岗位交接/公司/{code}.json').read_text(encoding='utf-8'))

    def save_doc(self, doc):
        (self.root / f"岗位交接/公司/{doc['公司编码']}.json").write_text(json.dumps(doc, ensure_ascii=False), encoding='utf-8')

    def master(self):
        return json.loads((self.root / '岗位交接/工作簿全量.json').read_text(encoding='utf-8'))

    def set_order(self, *pairs, ok=True):
        return self.run_cmd('set-order', '--input', self.write('order.json', [{"code": c, "志愿顺序": n} for c, n in pairs]),
                            '--quote', '游戏产品第一，AI第二', ok=ok)

    def handed(self):
        """甲公司两个岗位交接并确认。"""
        self.set_order((J_AI, 1), (J_GAME, 2))
        self.run_cmd('forward')
        self.run_cmd('forward-confirm', '--quote', '确认交接')

    def outcome(self, code, event, title):
        f = self.root / '岗位交接/投递回传.json'
        d = json.loads(f.read_text(encoding='utf-8')) if f.exists() else {'schemaVersion': 2, 'roles': {}, 'observations': []}
        d['roles'][code] = {'event': event, 'roleTitle': title, 'batchId': 'B'}
        f.write_text(json.dumps(d, ensure_ascii=False), encoding='utf-8')

    def observe(self, *items):
        f = self.root / '岗位交接/投递回传.json'
        d = json.loads(f.read_text(encoding='utf-8')) if f.exists() else {'schemaVersion': 2, 'roles': {}, 'observations': []}
        d['observations'] += [dict(o, evidence='测试', observedAt='2026-10-27T00:00:00Z') for o in items]
        f.write_text(json.dumps(d, ensure_ascii=False), encoding='utf-8')

    # ---------- 与投递组件的内容指纹一致 ----------
    def test_content_sha_matches_the_delivery_component_byte_for_byte(self):
        # 同一份宽表在投递组件 test/company-directory.test.mjs 里由 contentSha 算得同一个值
        doc = json.loads('{"schemaVersion": 3, "公司编码": "C00007", "公司名称": "网易", "志愿分组依据": null, "生成自": "test", "站点": [{"站点编码": "S0003", "站点": "https://campus.163.com", "登录链接": null, "账号核对页": null, "状态": "在用"}, {"站点编码": "S0004", "站点": "https://campus.game.163.com", "登录链接": "https://campus.game.163.com/login", "账号核对页": null, "状态": "在用"}, {"站点编码": "S0005", "站点": "https://leihuo.163.com", "登录链接": null, "账号核对页": null, "状态": "在用"}], "岗位": [{"岗位编码": "J000011", "地区": "中国大陆", "站点编码": "S0004", "岗位名称": "游戏产品经理", "岗位链接": "https://campus.game.163.com/job/4791", "投递链接": null, "官方岗位ID": "4791", "BU": null, "志愿分组": "公司", "志愿顺序": 2, "城市选项": ["杭州", "广州"], "投递状态": "岗位待投递", "链接说明": null, "材料": {"schemaVersion": 2, "items": [{"key": "referralCode", "status": "available", "value": "SampleCode17", "raw": "SampleCode17", "requireWhenAvailable": true, "evidence": {"岗位编码": "J000011", "column": "内推码", "cellSha256": "01ea865b8857c860a984e96930719596dbf5b4c002f7b3d00bb80d2da84b167e"}}]}}, {"岗位编码": "J000012", "地区": "中国大陆", "站点编码": "S0003", "岗位名称": "平台产品经理 “AI”", "岗位链接": "https://campus.163.com/job/88", "投递链接": null, "官方岗位ID": "88", "BU": null, "志愿分组": "公司", "志愿顺序": 1, "城市选项": [], "投递状态": "岗位待投递", "链接说明": null, "材料": {"schemaVersion": 2, "items": [{"key": "referralCode", "status": "available", "value": "SampleCode17", "raw": "SampleCode17", "requireWhenAvailable": true, "evidence": {"岗位编码": "J000012", "column": "内推码", "cellSha256": "01ea865b8857c860a984e96930719596dbf5b4c002f7b3d00bb80d2da84b167e"}}]}}, {"岗位编码": "J000013", "地区": "中国大陆", "站点编码": "S0005", "岗位名称": "美术", "岗位链接": "https://leihuo.163.com/j/1", "投递链接": null, "官方岗位ID": null, "BU": null, "志愿分组": "公司", "志愿顺序": null, "城市选项": [], "投递状态": "岗位待投递", "链接说明": null, "材料": {"schemaVersion": 2, "items": [{"key": "referralCode", "status": "available", "value": "SampleCode17", "raw": "SampleCode17", "requireWhenAvailable": true, "evidence": {"岗位编码": "J000013", "column": "内推码", "cellSha256": "01ea865b8857c860a984e96930719596dbf5b4c002f7b3d00bb80d2da84b167e"}}]}}], "批次": {"批次编号": "B1", "状态": "等待正向评审", "公司顺序": 1, "选中岗位": ["J000011", "J000012"], "排除": [], "执行": null, "确认": null, "进度": null, "更新时间": "t"}}')
        self.assertEqual(C.content_sha(doc), '50bf6d2b79a12e94ac74c14fab716476814c507aecf1e48429b3103865a781e6')

    # ---------- 编码迁移首次写 Excel ----------
    def test_code_migration_first_write_requires_identical_statuses_and_runs_once(self):
        ctx = S.Ctx(self.root, '产品管理/labels.json')
        with self.assertRaises(SystemExit):  # Excel 还没有编码列：正常门禁按编码配不上行
            S.write(ctx)
        m = self.master()
        sh = m['sheets'][0]
        idx = S.headers(sh)
        row = next(r for r in sh['rows'] if S.cell_value(r, idx['岗位名称']) == 'AI产品经理')
        cell = next(c for c in row['cells'] if c['column'] == idx['投递状态'])
        cell['value'] = '岗位已投递'
        (self.root / '岗位交接/工作簿全量.json').write_text(json.dumps(m, ensure_ascii=False), encoding='utf-8')
        with self.assertRaises(SystemExit) as e:
            S.write(ctx, code_migration=True)
        self.assertIn('逐行一致', str(e.exception))
        cell['value'] = '岗位待投递'
        S.refresh_digest(m)
        (self.root / '岗位交接/工作簿全量.json').write_text(json.dumps(m, ensure_ascii=False), encoding='utf-8')
        S.write(ctx, code_migration=True)
        from openpyxl import load_workbook
        head = [c.value for c in load_workbook(self.root / 'book.xlsx')['中国大陆'][1]]
        self.assertIn('岗位编码', head)
        with self.assertRaises(SystemExit) as e:
            S.write(ctx, code_migration=True)
        self.assertIn('迁移已完成', str(e.exception))
        S.write(ctx)

    # ---------- 读取与投影 ----------
    def test_role_id_is_a_value_read_from_the_link(self):
        self.assertEqual(C.Store.role_id('https://talent.baidu.com/jobs/detail/GRADUATE/0aa80bfa-a052-4052-afc1-28c38407c13c'), '0aa80bfa-a052-4052-afc1-28c38407c13c')
        self.assertEqual(C.Store.role_id('https://x.cn/pb/posDetail.html?postId=6a83c54768cc&postType=campus'), '6a83c54768cc')
        self.assertIsNone(C.Store.role_id('https://sensetime.jobs.feishu.cn'))
        self.assertEqual(C.Store.role_id('https://www.lixiang.com/employ/detail/19719.html?jobCode=A98366'), 'A98366')

    def test_refresh_writes_one_wide_file_per_company_from_the_master(self):
        out = self.run_cmd('refresh')
        self.assertEqual((out['written'], out['removed']), (3, []))
        self.assertEqual(sorted(p.name for p in (self.root / '岗位交接/公司').glob('*.json')), ['C00001.json', 'C00002.json', 'C00003.json'])
        jia = self.doc()
        self.assertEqual(set(jia), {'schemaVersion', '公司编码', '公司名称', '志愿分组依据', '站点', '岗位', '批次', '生成自'})
        self.assertEqual([(s['站点编码'], s['站点']) for s in jia['站点']], [('S0001', 'https://jia.jobs.example.cn'), ('S0002', 'https://game.jia.example.cn')])
        ai = jia['岗位'][0]
        self.assertEqual((ai['岗位编码'], ai['站点编码'], ai['岗位链接'], ai['官方岗位ID'], ai['投递链接'], ai['城市选项']),
                         (J_AI, 'S0001', AI, '123456', None, []))
        self.assertEqual(ai['材料']['items'][0]['evidence']['岗位编码'], J_AI)
        self.assertIsNone(jia['批次'])
        # 主表改了值：refresh 覆盖投影，编码不变
        self.sync('upsert', '--input', self.write('u.json', [{"key": {"code": J_AI}, "values": {"城市选项": ["北京", "上海"]}}]))
        self.run_cmd('refresh')
        self.assertEqual(self.doc()['岗位'][0]['城市选项'], ['北京', '上海'])

    def test_explicit_requirements_keep_both_roles_and_clear_only_after_batch(self):
        import application_requirements as R
        self.set_order((J_GAME, 1), (J_AI, 2))
        req = R.Requirements(self.root, 'test-session', 'test-turn')
        req.set({'公司': ['甲公司'], '浏览器': 'iab'}, '只做甲公司，两个岗位都要，只用内置浏览器')
        original = (self.root / '岗位交接/工作簿全量.json').read_bytes()
        out = self.run_cmd('forward')
        self.assertEqual((out['companies'], out['roles'], out['excluded']), (1, 2, 0))
        self.assertTrue(out['requirementsCleared'])
        b = self.doc()['批次']
        self.assertEqual((b['状态'], b['选中岗位']), ('等待正向评审', [J_GAME, J_AI]))
        self.assertNotIn('执行',b)
        self.assertEqual(json.loads((self.root/'.runtime/application/runs/batch-browser-policy.json').read_text())['browserChannels'],['iab'])
        self.assertEqual(json.loads(req.temp.read_text(encoding='utf-8')), {})
        self.assertEqual((self.root / '岗位交接/工作簿全量.json').read_bytes(), original)

    def test_invalid_explicit_company_does_not_consume_or_replace_batch(self):
        import application_requirements as R
        self.run_cmd('forward')
        before = {p.name: p.read_bytes() for p in (self.root / '岗位交接/公司').glob('*.json')}
        req = R.Requirements(self.root, 'test-session', 'test-turn')
        req.set({'公司': ['不存在的公司']}, '只做指定公司')
        result = self.run_cmd('forward', ok=False)
        self.assertIn('指定公司没有待投递岗位', result.stderr)
        self.assertEqual({p.name: p.read_bytes() for p in (self.root / '岗位交接/公司').glob('*.json')}, before)
        self.assertEqual(json.loads(req.temp.read_text(encoding='utf-8'))['requirements']['公司'], ['不存在的公司'])

    # ---------- 志愿顺序 ----------
    def test_forward_requires_user_given_order_within_each_group(self):
        out = self.run_cmd('forward')
        self.assertEqual(out['roles'], 0)
        reasons = {e['岗位编码']: e['原因'] for e in self.doc()['批次']['排除']}
        self.assertIn('志愿顺序未由用户给全', reasons[J_AI])
        self.assertNotEqual(self.set_order((J_GAME, 1), (J_AI, 1), ok=False).returncode, 0)
        self.assertNotEqual(self.set_order((J_GAME, 0), (J_AI, 3), ok=False).returncode, 0)
        self.set_order((J_GAME, 1), (J_AI, 2))
        ledger = json.loads((self.root / '产品管理/确认.json').read_text(encoding='utf-8'))
        self.assertEqual({(e['code'], e['value']) for e in ledger['value_confirmations'] if e['field'] == '志愿顺序'}, {(J_GAME, 1), (J_AI, 2)})
        self.run_cmd('forward')
        jia = self.doc()
        by = {r['岗位编码']: r for r in jia['岗位']}
        self.assertEqual([(by[j]['岗位名称'], by[j]['志愿顺序'], by[j]['志愿分组']) for j in jia['批次']['选中岗位']],
                         [('游戏产品', 1, '公司'), ('AI产品经理', 2, '公司')])

    def test_bu_parallel_quota_splits_order_groups(self):
        m = self.master()
        sh = m['sheets'][0]
        for r in sh['rows']:
            if r['row'] in (3, 4):  # 第2行是甲公司的公司行，岗位从第3行开始
                r['cells'].append({'column': 8, 'value': '甲BU' if r['row'] == 3 else '乙BU', 'type': 's'})
                r['cells'].sort(key=lambda c: c['column'])
        head = ['公司', '范围', '适用计划/批次', '投递上限', '当前投递数量', '是否分BU可平行投递', '是否可以滚动投递', '规则原文', '来源/核验日期', '公司编码']
        vals = ['甲公司', '公司总量', '全部计划', 2, 0, '是', '否', '原文', '来源：甲公司-FAQ 核验日期：2026-10-20', 'C00001']
        m['sheets'].append({'name': '投递名额', 'rows': [{'row': 1, 'cells': [{'column': i + 1, 'value': h, 'type': 's'} for i, h in enumerate(head)]},
                            {'row': 2, 'cells': [{'column': i + 1, 'value': v, 'type': 's' if isinstance(v, str) else 'n'} for i, v in enumerate(vals)]}],
                            'max_row': 2, 'max_column': len(head), 'state': 'visible', 'merged_ranges': [], 'tables': []})
        (self.root / '岗位交接/工作簿全量.json').write_text(json.dumps(m, ensure_ascii=False), encoding='utf-8')
        self.run_cmd('refresh')
        self.assertEqual([r['志愿分组'] for r in self.doc()['岗位']], ['甲BU', '乙BU'])
        self.assertEqual(self.run_cmd('forward')['roles'], 2)  # 每个分组只有一个岗位，不需要顺序

    # ---------- 正向评审 ----------
    def test_forward_rejects_non_role_links_and_confirms_per_company(self):
        self.set_order((J_AI, 1), (J_GAME, 2))
        out = self.run_cmd('forward')
        self.assertEqual(out['roles'], 2)
        self.assertIn('第三方', self.doc('C00002')['批次']['排除'][0]['原因'])
        reasons = {e['岗位']: e['原因'] for e in self.doc('C00003')['批次']['排除']}
        self.assertIn('共用同一链接', reasons['A岗'])
        self.assertIn('没有官方岗位ID', reasons['B岗'])
        self.assertEqual(self.doc('C00003')['批次']['状态'], '不交接')
        self.assertEqual(self.run_cmd('batch-status')['companies']['C00001']['handoffProblem'], '投递批次未经用户正向评审确认')
        self.run_cmd('forward-confirm', '--quote', '确认交接')
        self.assertIsNone(self.run_cmd('batch-status')['companies']['C00001']['handoffProblem'])
        jia = self.doc()
        self.assertEqual((jia['批次']['状态'], jia['批次']['确认']['userQuote']), ('已交接', '确认交接'))
        jia['岗位'][0]['岗位链接'] = 'https://jia.jobs.example.cn/tampered'
        self.save_doc(jia)
        self.assertIn('被改动', self.run_cmd('batch-status')['companies']['C00001']['handoffProblem'])
        self.assertNotEqual(self.run_cmd('forward', ok=False).returncode, 0)

    def test_direct_submission_confirmation_refreshes_handoff_progress(self):
        self.handed()
        self.sync('confirm', '--input', self.write('c.json', [{"code": J_AI, "to": "岗位已投递"}]), '--quote', '确认已投递')
        doc = self.doc()
        self.assertEqual(next(r['投递状态'] for r in doc['岗位'] if r['岗位编码'] == J_AI), '岗位已投递')
        self.assertIn(J_AI, doc['批次']['进度']['已处理'])
        self.assertNotIn(J_AI, doc['批次']['进度']['待投递'])

    def test_status_changes_elsewhere_do_not_break_the_confirmation(self):
        """内容指纹只覆盖执行内容：主表里改投递状态、JD 等不让已确认的交接失效。"""
        self.handed()
        self.sync('upsert', '--input', self.write('u.json', [{"key": {"code": J_AI}, "values": {"岗位JD": "新的JD"}}]))
        self.sync('confirm', '--input', self.write('c.json', [{"code": "J000003", "to": "岗位失效/关闭"}]), '--quote', '乙那个关了')
        self.run_cmd('refresh')
        self.assertIsNone(C.verify_handoff(self.doc()))

    def test_link_exception_lets_list_page_link_through_forward(self):
        self.set_order((J_GAME, 1), (J_AI, 2))
        lst = 'https://jia.jobs.example.cn/campus-recruit-list.html?positionId=123456'
        self.sync('relink', '--input', self.write('in.json', [{"code": J_AI, "new_link": lst}]), '--quote', '换成官网列表链接')
        self.run_cmd('forward')
        self.assertEqual({e['岗位编码']: e['原因'] for e in self.doc()['批次']['排除']}[J_AI], C.LIST_PAGE_PROBLEM)
        self.sync('link-exception', '--input', self.write('in.json', [{"code": J_AI, "定位说明": "列表页按 positionId 定位，点岗位名展开 JD"}]),
                  '--quote', '官网没有详情页，这个可以')
        result = self.run_cmd('forward')
        jia = self.doc()
        by = {r['岗位编码']: r for r in jia['岗位']}
        self.assertEqual([(by[j]['岗位名称'], by[j]['官方岗位ID'], by[j]['链接说明']) for j in jia['批次']['选中岗位']],
                         [('游戏产品', '778899', None), ('AI产品经理', '123456', '列表页按 positionId 定位，点岗位名展开 JD')])
        roles = result['handoffCompanies'][0]['岗位']
        self.assertEqual([r['岗位编码'] for r in roles], jia['批次']['选中岗位'])
        self.assertIn('列表页按 positionId 定位', next(r['链接说明'] for r in roles if r['岗位编码'] == J_AI))
        self.assertFalse((self.root / '岗位交接/正向评审.html').exists())

    def test_confirmations_are_kept_per_scope(self):
        self.handed()
        preview = self.run_cmd('amend-materials')
        self.assertEqual(preview['changes'], [])
        m = self.master()
        row = next(r for r in m['sheets'][0]['rows'] if r['row'] == 3)
        S.set_cell(row, 12, 'ABC123')  # 内推码
        S.refresh_digest(m)
        (self.root / '岗位交接/工作簿全量.json').write_text(json.dumps(m, ensure_ascii=False), encoding='utf-8')
        preview = self.run_cmd('amend-materials')
        self.assertEqual(len(preview['changes']), 1)
        self.run_cmd('amend-materials', '--expected-sha', preview['previousContentSha256'], '--quote', '补上推荐码')
        c = self.doc()['批次']['确认']
        self.assertEqual((c['userQuote'], c['scope']), ('补上推荐码', 'materials_only'))
        self.assertEqual(c['scopes']['selection']['userQuote'], '确认交接')
        self.assertEqual(set(c['scopes']), {'selection', 'materials'})
        self.assertIsNone(C.verify_handoff(self.doc()))

    # ---------- 链接观察 ----------
    def test_relink_of_a_role_in_a_running_batch_must_go_through_apply_observations(self):
        self.handed()
        p = self.sync('relink', '--input', self.write('r.json', [{"code": J_AI, "new_link": AI + '?v=2'}]), '--preview', ok=False)
        self.assertIn('apply-observations', p.stderr)

    def test_application_link_observation_is_confirmed_into_the_master_and_reconfirms_only_that_company(self):
        self.handed()
        self.sync('upsert', '--input', self.write('u.json', [{"key": {"sheet": "中国大陆", "company": "丙公司", "role": "new"},
                                                              "values": {"岗位名称": "C岗", "岗位链接": "https://bing.example.cn/job/9001"}}]))
        apply_url = 'https://jia.jobs.example.cn/form?jobAdId=X-1'
        self.observe({'kind': 'role', 'code': J_AI, 'field': '投递链接', 'value': apply_url})
        preview = self.run_cmd('apply-observations')
        self.assertEqual(preview['show_user_exactly'][0], f'中国大陆｜甲公司｜AI产品经理（{J_AI}）：投递链接「空」→「{apply_url}」')
        self.assertIn('只需重新导入这一家', preview['show_user_exactly'][1])
        self.assertNotEqual(self.run_cmd('apply-observations', '--quote', '可以', '--expected-sha', 'wrong', ok=False).returncode, 0)
        self.run_cmd('apply-observations', '--quote', '同意', '--expected-sha', preview['previousContentSha256'])
        jia = self.doc()
        self.assertEqual(jia['岗位'][0]['投递链接'], apply_url)
        self.assertEqual(jia['批次']['确认']['scopes']['links']['userQuote'], '同意')
        self.assertEqual(jia['批次']['确认']['scopes']['selection']['userQuote'], '确认交接')
        self.assertIsNone(C.verify_handoff(jia))
        self.assertFalse((self.root / '岗位交接/投递回传.json').exists())

    def test_dead_role_link_waits_for_a_new_link_then_keeps_the_code(self):
        self.handed()
        self.observe({'kind': 'role', 'code': J_AI, 'field': '岗位链接', 'value': None})
        preview = self.run_cmd('apply-observations')
        self.assertEqual(preview['show_user_exactly'], [])
        self.assertIn('岗位链接打不开', preview['waiting_for_new_link'][0])
        new = 'https://jia.jobs.example.cn/job/detail?jobAdId=G-777'
        repl = self.write('repl.json', [{"code": J_AI, "new_link": new}])
        preview = self.run_cmd('apply-observations', '--input', repl)
        self.run_cmd('apply-observations', '--input', repl, '--quote', '新链接对', '--expected-sha', preview['previousContentSha256'])
        jia = self.doc()
        ai = jia['岗位'][0]
        self.assertEqual((ai['岗位编码'], ai['岗位链接'], ai['官方岗位ID'], ai['志愿顺序']), (J_AI, new, 'G-777', 1))
        self.assertIsNone(C.verify_handoff(jia))
        self.assertFalse((self.root / '岗位交接/投递回传.json').exists())

    def test_redirected_role_link_with_a_new_id_is_previewed_and_replaces_the_link_after_confirmation(self):
        self.handed()
        landed = 'https://jia.jobs.example.cn/job/detail?jobAdId=739bc546-7ae3-4f6a-9d4f-681f02bf3a91'
        self.observe({'kind': 'role', 'code': J_AI, 'field': '岗位链接', 'value': landed,
                      'evidence': '岗位链接跳转后岗位编号由 1 变为 739bc546-7ae3-4f6a-9d4f-681f02bf3a91；页面岗位名与确认岗位名一致'})
        preview = self.run_cmd('apply-observations')
        self.assertEqual(preview['waiting_for_new_link'], [])
        self.assertIn(landed, preview['show_user_exactly'][0])
        self.run_cmd('apply-observations', '--quote', '换成跳转后的链接', '--expected-sha', preview['previousContentSha256'])
        ai = self.doc()['岗位'][0]
        self.assertEqual((ai['岗位编码'], ai['岗位链接'], ai['官方岗位ID']), (J_AI, landed, '739bc546-7ae3-4f6a-9d4f-681f02bf3a91'))
        self.assertFalse((self.root / '岗位交接/投递回传.json').exists())

    def test_site_link_observation_goes_to_the_site_row(self):
        self.run_cmd('refresh')
        self.observe({'kind': 'site', 'code': 'S0001', 'field': '账号核对页', 'value': 'https://jia.jobs.example.cn/me'},
                     {'kind': 'site', 'code': 'S0001', 'field': '登录链接', 'value': None})
        preview = self.run_cmd('apply-observations')
        self.assertEqual(preview['show_user_exactly'], ['甲公司｜站点 https://jia.jobs.example.cn（S0001）：账号核对页「空」→「https://jia.jobs.example.cn/me」'])
        self.run_cmd('apply-observations', '--quote', '对')
        self.assertEqual(self.doc()['站点'][0]['账号核对页'], 'https://jia.jobs.example.cn/me')

    # ---------- 反向评审 ----------
    def test_reverse_review_writes_master_only_after_user_confirmation(self):
        self.set_order((J_AI, 1), (J_GAME, 2))
        self.run_cmd('forward')
        self.assertNotEqual(self.run_cmd('reverse', ok=False).returncode, 0)  # 未确认交接不能反向评审
        self.run_cmd('forward-confirm', '--quote', '确认交接')
        self.outcome(J_AI, 'submission_confirmed', 'AI产品经理')
        self.outcome(J_GAME, 'submission_uncertain', '游戏产品')
        out = self.run_cmd('reverse')
        self.assertEqual(out['show_user_exactly'], [f'中国大陆｜甲公司｜AI产品经理（{J_AI}）：当前状态是「岗位待投递」，更新状态是「岗位已投递」'])
        self.assertIn('游戏产品：投递结果不确定或受阻，处理前禁止重跑', out['notes'])
        self.assertNotIn('岗位已投递', (self.root / '岗位交接/工作簿全量.json').read_text(encoding='utf-8'))
        self.assertNotEqual(self.run_cmd('forward', ok=False).returncode, 0)  # 未回写完不能开下一批
        res = self.run_cmd('reverse-confirm', '--quote', '确认回写')
        self.assertEqual((res['companies'], res['statusChanges'], res['unresolved']), ({'C00001': '等待反向评审'}, 1, [J_GAME]))
        self.assertIn('岗位已投递', (self.root / '岗位交接/工作簿全量.json').read_text(encoding='utf-8'))
        self.run_cmd('reverse')
        res = self.run_cmd('reverse-confirm', '--quote', '不确定那个我已核实没投上', '--resolved', J_GAME)
        self.assertEqual(res['companies'], {'C00001': '已回写'})
        self.assertFalse((self.root / '岗位交接/投递回传.json').exists())
        self.run_cmd('forward')  # 本轮回写完才能开下一轮

    def test_partial_reverse_keeps_original_confirmed_batch_runnable(self):
        self.handed()
        before = C.content_sha(self.doc())
        self.outcome(J_AI, 'submission_confirmed', 'AI产品经理')
        self.run_cmd('reverse')
        got = self.run_cmd('reverse-confirm', '--quote', '确认这一条')
        self.assertEqual(got['companies'], {'C00001': '进行中'})
        jia = self.doc()
        self.assertEqual(jia['批次']['进度'], {'已处理': [J_AI], '待投递': [J_GAME]})
        self.assertEqual(C.content_sha(jia), before)
        self.assertIsNone(C.verify_handoff(jia))
        self.assertNotEqual(self.run_cmd('reverse-confirm', '--quote', '再确认', ok=False).returncode, 0)
        self.outcome(J_GAME, 'submission_confirmed', '游戏产品')
        self.run_cmd('reverse')
        self.assertEqual(self.run_cmd('reverse-confirm', '--quote', '确认最后一条')['companies'], {'C00001': '已回写'})

    def test_reverse_confirm_keeps_pending_link_observations(self):
        self.handed()
        self.outcome(J_AI, 'submission_confirmed', 'AI产品经理')
        self.observe({'kind': 'role', 'code': J_GAME, 'field': '投递链接', 'value': 'https://game.jia.example.cn/apply/1'})
        self.run_cmd('reverse')
        self.run_cmd('reverse-confirm', '--quote', '确认这一条')
        d = json.loads((self.root / '岗位交接/投递回传.json').read_text(encoding='utf-8'))
        self.assertEqual([o['code'] for o in d['observations']], [J_GAME])


if __name__ == '__main__':
    unittest.main()
