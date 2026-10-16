import copy
import unittest
import company_handoff as h
from user_priorities import validate
from test_company_handoff import CompanyHandoffTests


class UserPriorityTests(CompanyHandoffTests):
    # Inherit integration coverage under the new mandatory user-order contract.
    def test_missing_user_rank_blocks_confirm(self):
        self.c['user_priorities']=[]
        self.draft()
        self.assertIn('已选岗位缺用户确认的公司或BU志愿顺序: r1',h.validate_research(h.read(self.book,self.batch)[0],True))

    def test_bu_independent_first_choices_roundtrip(self):
        self.c['priority_scope']='bu'
        for p in self.c['user_priorities']:
            role=next(r for r in self.c['roles'] if r['id']==p['role_id'])
            p.update(scope='bu',bu=role['bu'],group_id=self.c['company_key']+'::'+role['bu'],rank=1)
        self.approve();c=h.read(self.book,self.batch)[0]
        self.assertEqual([p['rank'] for p in c['user_priorities']],[1,1]);self.assertEqual(h.check(self.book,self.batch),[])

    def test_company_duplicate_rank_and_wrong_bu(self):
        c=copy.deepcopy(self.c);c['user_priorities'][1]['rank']=1
        self.assertIn('同一用户志愿组序号重复',validate(c))
        c=copy.deepcopy(self.c);c['priority_scope']='bu'
        for p in c['user_priorities']:p.update(scope='bu',bu='wrong',group_id='synthetic::wrong')
        self.assertIn('BU内志愿必须对应岗位的已核实BU',validate(c))

    def test_portal_parallel_preserves_separate_user_ranks(self):
        for p in self.c['preferences']:p.update(mode='parallel',status='pending_selection',rank=None,group_id='parallel',label='parallel')
        next(r for r in self.c['rules'] if r['dimension']=='preference')['preference_mode']='parallel'
        self.c['roles'][0]['selected']=False
        self.draft();c=h.read(self.book,self.batch)[0]
        self.assertEqual([p['rank'] for p in c['preferences']],[None,None])
        self.assertEqual([p['rank'] for p in c['user_priorities']],[1,2])

    def test_user_order_evidence_tamper_blocks(self):
        self.approve();(self.root/'priority-confirm.md').write_text('changed',encoding='utf-8')
        self.assertIn('用户志愿确认依据已变化',h.check(self.book,self.batch))
