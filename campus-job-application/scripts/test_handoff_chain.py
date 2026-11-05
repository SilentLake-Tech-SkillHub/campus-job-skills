"""全链路测试：主表（固定编码）→ 每家公司宽表 → 正向评审 → 投递组件读取已确认的公司文件 →
投递组件记链接观察 → apply-observations 用户确认后写主表、只重确认这一家 → 投递组件写回传 → 反向评审 → 用户确认写回主表。

投递组件（job-application-assistant-v2）不在同一仓库；不存在或没有 node 时跳过。
"""
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
import uuid
from pathlib import Path

from openpyxl import Workbook

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import workbook_store  # noqa: E402
import json_excel_sync as S  # noqa: E402
import master_codes as M  # noqa: E402

V2 = HERE.parents[1] / 'job-application-assistant-v2'
STORE = HERE / 'company_store.py'
SYNC = HERE / 'json_excel_sync.py'
HEAD = ['地区', '公司', '主要行业', '企业性质', '岗位名称', '工作地点', '岗位JD', '所属BU', '岗位链接', '投递状态', '备注']
A_URL = 'https://jia.jobs.example.cn/campus/position/123456/detail'
B_URL = 'https://jia.jobs.example.cn/campus/position/778899/detail'

NODE = r'''
import {readFile, writeFile} from 'node:fs/promises';
import {createHash, randomUUID} from 'node:crypto';
import path from 'node:path';
const v2 = process.env.V2_DIR;
const {readHandoffBatch} = await import(path.join(v2, 'shared/company-directory.mjs'));
const {closeApplication} = await import(path.join(v2, 'shared/application-closeout.mjs'));
const {recordObservation} = await import(path.join(v2, 'shared/application-outcome.mjs'));
const sha = b => createHash('sha256').update(b).digest('hex');
const {input} = await readHandoffBatch();
const results = [];
// 链接观察：投递端只记到 投递回传.json，不改交接区
for (const o of JSON.parse(process.env.OBSERVE || '[]')) results.push(await recordObservation({...o, evidence: 'role_enter_application_verified'}));
for (const [title, event, result] of JSON.parse(process.env.PLAN)) {
  const c = input.companies.find(x => x.roles.some(r => r.title === title));
  const r = c.roles.find(x => x.title === title);
  const shot = path.join(process.env.EVIDENCE_DIR, r.roleId + '.png');
  await writeFile(shot, 'png-' + r.roleId);
  const ev = {schemaVersion: 1, event, batchId: input.batchId, companyId: r.siteId, roleId: r.roleId, roleUrl: r.url, roleTitle: r.title,
    observedAt: new Date(Date.now() - 1000).toISOString(), sourceUrl: r.url, result, reviewer: 'test',
    artifacts: [{file: shot, sha256: sha(await readFile(shot))}]};
  const evFile = path.join(process.env.EVIDENCE_DIR, r.roleId + '.json');
  await writeFile(evFile, JSON.stringify(ev));
  results.push(await closeApplication({outcome: {schemaVersion: 1, eventId: randomUUID(), event, batchId: input.batchId,
    companyId: r.siteId, roleId: r.roleId, roleUrl: r.url, roleTitle: r.title,
    evidence: {file: evFile, sha256: sha(await readFile(evFile))}}}));
}
console.log(JSON.stringify({input, results}));
'''


@unittest.skipUnless((V2 / 'shared/company-directory.mjs').exists() and shutil.which('node'), '投递组件或 node 不可用')
class HandoffChainTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        r = self.root = Path(self.tmp.name)
        (r / '产品管理').mkdir()
        (r / '岗位交接').mkdir()
        (r / 'evidence').mkdir()
        wb = Workbook()
        ws = wb.active
        ws.title = '中国大陆'
        ws.append(HEAD)
        ws.append(['上海', '甲公司', '互联网', '私企', None, None, None, None, None, '已有岗位', None])
        ws.append(['上海', '甲公司', '互联网', '私企', 'AI产品经理', '上海', 'JD', '', A_URL, '岗位待投递', None])
        ws.append(['上海', '甲公司', '互联网', '私企', '数据产品经理', '上海', 'JD', '', B_URL, '岗位待投递', None])
        ws.append(['北京', '乙公司', '互联网', '私企', None, None, None, None, None, '待搜索', None])
        wb.save(r / 'book.xlsx')
        (r / '产品管理/company-labels.json').write_text(json.dumps({"schemaVersion": 1, "rowKind": "company",
            "labels": [{"value": "待搜索"}, {"value": "已有岗位"}]}, ensure_ascii=False), encoding='utf-8')
        labels = [{"value": "岗位待投递", "definition": "x"},
                  {"value": "岗位已投递", "definition": "x"},
                  {"value": "岗位失效/关闭", "definition": "x"}]
        cfg = {"schemaVersion": 1, "rowKind": "role", "labelGroups": ["产品管理/company-labels.json"], "column": "投递状态", "sheets": ["中国大陆"], "labels": labels,
               "confirmationLedger": "产品管理/确认.json", "masterData": "岗位交接/工作簿全量.json", "workbook": "book.xlsx",
               "companyStore": {"dir": "岗位交接/公司", "forwardReview": "岗位交接/正向评审.html",
                                "thirdPartyHosts": []}}
        (r / '产品管理/labels.json').write_text(json.dumps(cfg, ensure_ascii=False), encoding='utf-8')
        # 迁移后的主表：甲公司 C00001、乙公司 C00002；AI产品经理 J000001、数据产品经理 J000002，站点 S0001
        snap = workbook_store.snapshot(r / 'book.xlsx')
        M.assign_missing_codes(snap, ['中国大陆'])
        S.refresh_digest(snap)
        (r / '岗位交接/工作簿全量.json').write_text(json.dumps(snap, ensure_ascii=False), encoding='utf-8')
        # 迁移脚本的首次写入：Excel 加上编码列（状态逐行一致才允许）
        S.write(S.Ctx(r, r / '产品管理/labels.json'), code_migration=True)

    def tearDown(self):
        self.tmp.cleanup()

    def py(self, script, *args):
        p = subprocess.run([sys.executable, '-B', str(script), '--workspace', str(self.root), '--labels', '产品管理/labels.json', *args],
                           capture_output=True, text=True, encoding='utf-8',
                           env={**os.environ, 'CODEX_THREAD_ID': 'chain-session', 'CODEX_TURN_ID': 'chain-turn', 'DSH_SESSION_ID': ''})
        if p.returncode:
            self.fail(p.stderr + p.stdout)
        return json.loads(p.stdout) if p.stdout.strip() else None

    def v2(self, plan, observe=()):
        env = dict(os.environ, V2_DIR=str(V2), RECRUITING_HANDOFF_DIR=str(self.root / '岗位交接'),
                   EVIDENCE_DIR=str(self.root / 'evidence'), PLAN=json.dumps(plan, ensure_ascii=False),
                   OBSERVE=json.dumps([dict(o, handoffDir=str(self.root / '岗位交接')) for o in observe], ensure_ascii=False))
        p = subprocess.run(['node', '--input-type=module', '-e', NODE], capture_output=True, text=True, encoding='utf-8', env=env, cwd=str(V2))
        if p.returncode:
            self.fail(p.stderr + p.stdout)
        return json.loads(p.stdout)

    def status(self):
        m = json.loads((self.root / '岗位交接/工作簿全量.json').read_text(encoding='utf-8'))
        rows = m['sheets'][0]['rows']
        return {next(c['value'] for c in r['cells'] if c['column'] == 5): next((c['value'] for c in r['cells'] if c['column'] == 10), None)
                for r in rows[1:] if any(c['column'] == 5 for c in r['cells'])}

    def test_full_chain(self):
        self.py(STORE, 'refresh')
        jia = json.loads((self.root / '岗位交接/公司/C00001.json').read_text(encoding='utf-8'))
        self.assertEqual([(r['岗位编码'], r['站点编码'], r['官方岗位ID']) for r in jia['岗位']], [('J000001', 'S0001', '123456'), ('J000002', 'S0001', '778899')])
        self.assertEqual(self.py(STORE, 'forward')['roles'], 0)  # 同组两个岗位，用户未给志愿顺序不交接
        order = self.root / 'order.json'
        order.write_text(json.dumps([{"code": "J000001", "志愿顺序": 1}, {"code": "J000002", "志愿顺序": 2}], ensure_ascii=False), encoding='utf-8')
        self.py(STORE, 'set-order', '--input', str(order), '--quote', 'AI产品第一志愿，数据产品第二')
        self.assertEqual(self.py(STORE, 'forward')['roles'], 2)
        # 正向评审未确认：投递组件拒绝读取
        env = dict(os.environ, V2_DIR=str(V2), RECRUITING_HANDOFF_DIR=str(self.root / '岗位交接'),
                   EVIDENCE_DIR=str(self.root / 'evidence'), PLAN='[]')
        p = subprocess.run(['node', '--input-type=module', '-e', NODE], capture_output=True, text=True, encoding='utf-8', env=env, cwd=str(V2))
        self.assertNotEqual(p.returncode, 0)
        self.assertIn('HANDOFF_NOT_CONFIRMED', p.stderr)

        self.py(STORE, 'forward-confirm', '--quote', '确认交接这两个岗位')
        # 投递端第一次进入 AI 岗位的申请页（投递链接原为空）：只记观察
        apply_url = 'https://jia.jobs.example.cn/campus/apply/123456'
        before = (self.root / '岗位交接/公司/C00001.json').read_text(encoding='utf-8')
        confirmed_sha = json.loads(before)['批次']['确认']['contentSha256']
        out = self.v2([], observe=[{'kind': 'role', 'code': 'J000001', 'field': '投递链接', 'value': apply_url}])
        self.assertEqual(out['results'], [{'status': 'observed', 'code': 'J000001', 'field': '投递链接'}])
        self.assertEqual(out['input']['companies'][0]['roles'][0].get('applyUrl'), None)
        self.assertEqual((self.root / '岗位交接/公司/C00001.json').read_text(encoding='utf-8'), before)
        # 搜索端预览 → 用户原话确认 → 写主表、只重生成并重确认这一家（范围：链接变更）
        preview = self.py(STORE, 'apply-observations')
        self.assertEqual(preview['show_user_exactly'][0], f'中国大陆｜甲公司｜AI产品经理（J000001）：投递链接「空」→「{apply_url}」')
        self.py(STORE, 'apply-observations', '--quote', '申请页地址对', '--expected-sha', preview['previousContentSha256'])
        jia = json.loads((self.root / '岗位交接/公司/C00001.json').read_text(encoding='utf-8'))
        self.assertEqual((jia['批次']['确认']['scope'], jia['批次']['确认']['previousContentSha256']), ('links_only', confirmed_sha))

        out = self.v2([['AI产品经理', 'submission_confirmed', 'official_receipt_confirmed'],
                       ['数据产品经理', 'role_closed', 'official_role_closed']])
        roles = out['input']['companies'][0]['roles']
        self.assertEqual([(r['roleId'], r['siteId'], r['title']) for r in roles], [('J000001', 'S0001', 'AI产品经理'), ('J000002', 'S0001', '数据产品经理')])
        self.assertEqual(roles[0].get('applyUrl'), apply_url)
        self.assertEqual({r['status'] for r in out['results']}, {'closed'})
        back = json.loads((self.root / '岗位交接/投递回传.json').read_text(encoding='utf-8'))
        self.assertEqual(sorted(back['roles']), ['J000001', 'J000002'])
        self.assertEqual(self.status(), {'AI产品经理': '岗位待投递', '数据产品经理': '岗位待投递'})  # 投递组件不改主表

        rev = self.py(STORE, 'reverse')
        self.assertEqual(sorted(rev['show_user_exactly']), sorted([
            '中国大陆｜甲公司｜AI产品经理（J000001）：当前状态是「岗位待投递」，更新状态是「岗位已投递」',
            '中国大陆｜甲公司｜数据产品经理（J000002）：当前状态是「岗位待投递」，更新状态是「岗位失效/关闭」']))
        self.assertEqual(self.status()['AI产品经理'], '岗位待投递')  # 反向评审只出评审页

        res = self.py(STORE, 'reverse-confirm', '--quote', '确认回写')
        self.assertEqual((res['companies'], res['statusChanges'], res['unresolved']), ({'C00001': '已回写'}, 2, []))
        self.assertEqual(self.status(), {'AI产品经理': '岗位已投递', '数据产品经理': '岗位失效/关闭'})
        self.assertFalse((self.root / '岗位交接/投递回传.json').exists())
        # 主表与确认台账一致：写 Excel 的门禁不报未确认变更；投递链接写进主表 J 行
        chk = self.py(SYNC, 'check')
        self.assertEqual(chk['status_changes_without_user_confirmation'], 0)
        self.py(SYNC, 'write')


if __name__ == '__main__':
    unittest.main()
