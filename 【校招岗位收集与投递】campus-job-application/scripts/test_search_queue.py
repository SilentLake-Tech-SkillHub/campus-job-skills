import asyncio
import copy
from concurrent.futures import ThreadPoolExecutor
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from search_queue import Queue,CHECKS,sha


class SearchQueueTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.root=Path(self.tmp.name)
        (self.root/'authorization.md').write_text('Synthetic test authorization',encoding='utf-8')
        self.q=Queue(self.root)
        self.manifest={'batch_id':'synthetic','scope':{'region':'synthetic'},'authorization_ref':'authorization.md',
            'workbook_sha256':'a'*64,'companies':[{'company_key':str(i),'name':'Synthetic '+str(i),'source_key':str(i)} for i in range(3)]}
        self.q.create(self.manifest)

    def tearDown(self):self.tmp.cleanup()

    def evidence(self,a,**updates):
        ref=a['task_id']+'.txt';(self.root/ref).write_text('Synthetic browser DOM capture',encoding='utf-8')
        e={'task_id':a['task_id'],'company_key':a['company']['company_key'],'channel':a['channel'],'stage':a['stage'],
           'observed_at':'test-observation','observations':['Synthetic source identity and empty directory'],
           'checks':dict.fromkeys(CHECKS[1],True),'gaps':[], 'source_urls':['https://example.test/campus'],
           'artifacts':[{'path':ref,'sha256':sha(self.root/ref)}],
           'inventory':{'plans':[],'roles':[],'expected_plans':0,'expected_roles':0}}
        e.update(updates);name=a['task_id']+'.json';(self.root/name).write_text(json.dumps(e),encoding='utf-8');return name

    def record(self,a,**updates):
        return self.q.record('synthetic',a['task_id'],a['lease_token'],self.evidence(a,**updates))

    def test_two_concurrent_dispatchers_do_not_duplicate(self):
        with ThreadPoolExecutor(2) as pool:
            groups=list(pool.map(lambda _:Queue(self.root).dispatch('synthetic'),range(2)))
        jobs=[a for group in groups for a in group['assignments']]
        self.assertEqual(len(jobs),2);self.assertEqual(len({a['task_id'] for a in jobs}),2)
        self.assertEqual({a['channel'] for a in jobs},{'iab','chrome'})

    def test_shared_source_alias_is_exclusive(self):
        m=copy.deepcopy(self.manifest);m['batch_id']='aliases'
        for c in m['companies']:c['source_key']='shared-ats'
        self.q.create(m);self.assertEqual(len(self.q.dispatch('aliases')['assignments']),1)

    def test_blocker_does_not_stop_independent_channel(self):
        a,b=self.q.dispatch('synthetic')['assignments']
        self.assertEqual(self.record(a,gaps=['Unavailable source'])['status'],'blocked')
        c=self.q.dispatch('synthetic')['assignments'][0]
        self.assertEqual(c['channel'],a['channel']);self.assertNotEqual(c['task_id'],b['task_id'])
        with self.assertRaises(ValueError):self.q.advance('synthetic')

    def test_stale_and_duplicate_delivery_rejected(self):
        a=self.q.dispatch('synthetic')['assignments'][0];self.record(a)
        with self.assertRaises(ValueError):self.record(a)

    def test_expired_lease_no_automatic_retry(self):
        a=self.q.dispatch('synthetic',['chrome'],30)['assignments'][0]
        with patch('search_queue.time.time',return_value=a['expires']+1):
            with self.assertRaises(ValueError):self.record(a)
            self.q.dispatch('synthetic',['chrome'])
        self.assertEqual(self.q.check('synthetic')['counts']['blocked'],1)
        self.q.retry('synthetic',a['task_id'],'Original tab inspected and source recovered')

    def test_wrong_channel_and_count_mismatch(self):
        a=self.q.dispatch('synthetic')['assignments'][0]
        for extra in [{'channel':'wrong'},{'inventory':{'plans':[],'roles':[],'expected_plans':2,'expected_roles':0}}]:
            with self.assertRaises(ValueError):self.record(a,**extra)

    def test_round_complete_and_capture_tamper(self):
        while True:
            jobs=self.q.dispatch('synthetic')['assignments']
            if not jobs:break
            for a in jobs:self.record(a)
        self.assertTrue(self.q.check('synthetic')['passed'])
        artifact=next(self.root.glob('*.txt'));original=artifact.read_text();artifact.write_text('changed')
        self.assertFalse(self.q.check('synthetic')['passed'])
        with self.assertRaises(ValueError):self.q.advance('synthetic')
        artifact.write_text(original);self.assertEqual(self.q.advance('synthetic')['stage'],2)

    def test_scope_frozen(self):
        m=copy.deepcopy(self.manifest);m['scope']={'region':'changed'}
        with self.assertRaises(ValueError):self.q.create(m)

    def test_all_outcomes_checked_carry_blockers_without_claiming_pass(self):
        while True:
            jobs=self.q.dispatch('synthetic')['assignments']
            if not jobs:break
            for a in jobs:self.record(a,gaps=['Source does not disclose current cycle'])
        checked=self.q.check('synthetic')
        self.assertTrue(checked['audited']);self.assertFalse(checked['passed'])
        result=self.q.advance('synthetic')
        self.assertEqual(result['stage'],2);self.assertEqual(result['previous_round_blocked'],3)
        self.assertFalse(result['previous_round_passed'])

    def test_real_mcp_stdio_protocol(self):
        from search_mcp_client import call
        with patch.dict(os.environ,{'CAMPUS_SEARCH_PRIVATE_DIR':str(self.root)}):
            tools=asyncio.run(call({'method':'tools/list'}))
            self.assertIn('search_dispatch',{t['name'] for t in tools['tools']})
            result=asyncio.run(call({'tool':'search_dispatch','arguments':{'batch_id':'synthetic'}}))
            self.assertEqual(len(result['assignments']),2)
