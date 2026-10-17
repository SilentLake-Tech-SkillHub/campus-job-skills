import copy
import unittest
from company_summary import summarize, render_summary

class CompanySummaryTests(unittest.TestCase):
    def sample(self):
        return {'name':'示例公司<script>','scope':'本批产品岗范围','plans':[{'id':'p1','name':'普通校招','category':'regular'},{'id':'p2','name':'星辰管培','category':'management_trainee'}], 'roles':[{'id':'a','tier':'exact','employment':'正式','selected':True},{'id':'b','tier':'partial','employment':'实习','selected':False},{'id':'c','tier':'outside','employment':'正式','selected':False}], 'conditions':[{'subject_type':'role','subject_id':'b','field':'conversion_opportunity','condition_value':True,'status':'verified','source':'https://example.test/job/b','raw':'有转正机会'}], 'research':[{'id':k,'status':'checked','expected_count':n,'observed_count':n,'gap':''} for k,n in [('plans',2),('catalog',3),('jd',2),('eligibility',2),('rules',1),('referrals',1),('preferences',1)]]}
    def test_counts_and_duplicate_identity(self):
        c=self.sample();c['roles'].append(copy.deepcopy(c['roles'][0]));s=summarize(c)
        self.assertEqual((s['matching_roles']['count'],s['formal_roles']['count'],s['conversion_roles']['count'],s['selected_roles']),(2,1,1,1))
        self.assertEqual(s['trainee_plans']['names'],['星辰管培'])
    def test_report_not_published_before_whole_company_scan(self):
        for key in ('plans','catalog','jd','eligibility','rules','referrals','preferences'):
            c=self.sample()
            next(x for x in c['research'] if x['id']==key)['status']='todo'
            self.assertEqual(render_summary(c),'',key)
        self.assertIn('公司范围与岗位汇报',render_summary(self.sample()))
    def test_verified_zero(self):
        c=self.sample();c['plans']=[];c['roles']=[];c['conditions']=[]
        for x in c['research']:x.update(expected_count=0,observed_count=0)
        s=summarize(c)
        self.assertEqual(s['trainee_plans']['count'],0);self.assertEqual(s['conversion_roles']['count'],0)
    def test_incomplete_not_zero(self):
        c=self.sample();c['research'][0]['status']='todo';s=summarize(c)
        self.assertIsNone(s['trainee_plans']['count']);self.assertIsNone(s['matching_roles']['count'])
    def test_missing_conversion_evidence_not_zero(self):
        c=self.sample();c['conditions']=[]
        self.assertIsNone(summarize(c)['conversion_roles']['count'])
    def test_html_escape_and_conflict(self):
        c=self.sample();self.assertNotIn('<script>',render_summary(c))
        c['roles'].append({**c['roles'][0],'employment':'实习'})
        with self.assertRaises(ValueError):summarize(c)

if __name__=='__main__':unittest.main()
