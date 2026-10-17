import copy
import json
import unittest
from unittest.mock import patch
import company_handoff as h
from approved_review import business_digest, validate_binding
import test_company_handoff as fixtures

class ApprovedReviewTests(unittest.TestCase):
 setUp=fixtures.CompanyHandoffTests.setUp
 tearDown=fixtures.CompanyHandoffTests.tearDown
 def frozen(self):
  c=copy.deepcopy(self.c)
  (self.root/'final.html').write_text('<html>用户已核对的唯一页面</html>',encoding='utf-8')
  (self.root/'approval.txt').write_text('合成用户确认本版本',encoding='utf-8')
  for kind,(_,fields) in h.SCHEMA.items():
   for item in ([c] if kind=='company' else c[kind]):
    for k,_ in fields:item.setdefault(k,None if k in h.NULL_FIELDS|h.JSON_FIELDS else [] if k in h.LIST_FIELDS else False if k in h.BOOL_FIELDS else '')
  c.update(review_mode='frozen',review_html_ref='final.html',review_html_sha256=h.sha(self.root/'final.html'),review_evidence_ref='approval.txt',review_evidence_sha256=h.sha(self.root/'approval.txt'),reviewed_at='20XX-10-17')
  c['review_business_sha256']=business_digest(c)
  return c
 def test_fixed_write_refresh_preserves_html(self):
  c=self.frozen();original=(self.root/'final.html').read_bytes()
  with patch.object(h,'render',side_effect=AssertionError('不得重做已确认HTML')):
   h.write(self.book,self.batch,{'companies':[c]});h.export(self.book,self.batch)
  self.assertEqual((self.root/'final.html').read_bytes(),original)
  self.assertFalse((h.paths(self.book,self.batch)/'公司核对.html').exists())
  payload=json.loads((h.paths(self.book,self.batch)/'公司方案.json').read_text())
  self.assertEqual(payload['approved_reviews'][0]['review_html_ref'],'final.html')
  self.assertFalse(validate_binding(h.read(self.book,self.batch)[0],self.root))
 def test_changes_rejected_before_book_write(self):
  c=self.frozen();before=h.sha(self.book);c['roles'][0]['title']='被改动岗位'
  with self.assertRaisesRegex(ValueError,'方案不一致'):h.write(self.book,self.batch,{'companies':[c]})
  self.assertEqual(h.sha(self.book),before)
 def test_tampered_html_rejected(self):
  c=self.frozen();(self.root/'final.html').write_text('被改动')
  with self.assertRaisesRegex(ValueError,'HTML哈希不一致'):h.write(self.book,self.batch,{'companies':[c]})
 def test_explicit_decline_cannot_be_selected(self):
  c=self.frozen();c['roles'][0]['application_decision']='decline';c['roles'][0]['decision_ref']='approval.txt';c['roles'][0]['decision_sha256']=h.sha(self.root/'approval.txt');c['review_business_sha256']=business_digest(c)
  with self.assertRaisesRegex(ValueError,'不可设为已选'):h.write(self.book,self.batch,{'companies':[c]})

 def test_second_intent_is_not_a_second_application(self):
  c=self.frozen();second=c['roles'][1];second.update(application_decision='second_intent',decision_ref='approval.txt',decision_sha256=h.sha(self.root/'approval.txt'),selected=False,recommended=False)
  next(r for r in c['rules'] if r['dimension']=='company').update(limit=1)
  c['review_business_sha256']=business_digest(c)
  h.write(self.book,self.batch,{'companies':[c]})
  saved=h.read(self.book,self.batch)[0]
  self.assertEqual(sum(r['selected'] for r in saved['roles']),1)
  self.assertEqual(saved['roles'][1]['application_decision'],'second_intent')
 def test_second_intent_cannot_also_be_selected(self):
  c=self.frozen();c['roles'][0].update(application_decision='second_intent',decision_ref='approval.txt',decision_sha256=h.sha(self.root/'approval.txt'));c['review_business_sha256']=business_digest(c)
  with self.assertRaisesRegex(ValueError,'不可设为已选'):h.write(self.book,self.batch,{'companies':[c]})

if __name__=='__main__':unittest.main()
