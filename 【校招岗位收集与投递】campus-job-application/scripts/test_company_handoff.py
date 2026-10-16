"""Synthetic company review: no live company, account, application or user approval."""
import copy
import json
import tempfile
import unittest
from unittest.mock import patch
from pathlib import Path
from openpyxl import Workbook, load_workbook
import company_handoff as h


def sample_company(book,name='合成公司',key='synthetic'):
 wb=load_workbook(book);roles=[]
 for ws in wb:
  if ws.title in {v[0] for v in h.SCHEMA.values()}:continue
  heads=[x.value for x in ws[1]]
  if '岗位链接' not in heads:continue
  for values in ws.iter_rows(min_row=2,values_only=True):
   row=dict(zip(heads,values))
   if row.get('公司')!=name or not row.get('岗位链接'):continue
   roles.append({'id':'r'+str(len(roles)+1),'plan_ids':['regular'],'title':row['岗位名称'],'url':row['岗位链接'],'bu':row['所属BU'],'cities':row['工作地点'],'round_id':'current','city_options':[row['工作地点']],'employment':'正式','tier':'exact','functional_fit':'符合','jd':row['岗位JD'],'basis':'合成需求及JD匹配','eligibility':'符合','eligibility_basis':'合成材料','recommended':True,'selected':True,'order':len(roles)+1,'source':'https://example.test/jobs','checked_at':'20XX-10-16'})
 wb.close()
 return {'company_key':key,'name':name,'version':'v1','scope_status':'verified','regions':['合成地区'],'cohort':'测试届别','employment_types':['正式'],'role_families':['AI产品'],'hard_cities':[],'preferred_cities':[],'scope':'合成测试：AI产品、城市不限','coverage':'complete','gaps':'','checked_at':'20XX-10-16','recommendation_reason':'基于合成JD和名额建议','plans':[{'category':'regular','window_status':'not_disclosed','requirements_status':'structured','degree_levels':['本科'],'majors':[],'id':'regular','name':'普通校招','type':'校招','requirements':'毕业窗口与学历，见合成来源','window':'本招聘轮次','source':'https://example.test/campus','checked_at':'20XX-10-16','special':False,'decision':'consider','decision_ref':''}], 'roles':roles,'conditions':[],'preferences':[{'id':'p'+r['id'],'role_id':r['id'],'group_id':'','rank':None,'label':'','mode':'not_applicable','bu':'','cities':'','status':'not_applicable','source':'https://example.test/faq','checked_at':'20XX-10-16'} for r in roles],'research':[{'id':item,'status':'checked','source':'https://example.test/research','checked_at':'20XX-10-16','gap':'','expected_count':(1 if item=='plans' else len(roles)) if item in ('plans','catalog') else '', 'observed_count':(1 if item=='plans' else len(roles)) if item in ('plans','catalog') else ''} for item in sorted(h.RESEARCH_ITEMS)],'rules':[{'id':dim,'dimension':dim,'scope':'全部岗位','plan_ids':[],'bu_ids':[],'round_id':'current' if dim=='round' else '', 'limit_type':'count' if dim=='company' else 'unlimited' if dim not in ('cooldown','change') else 'not_applicable','preference_mode':'not_applicable','usage_status':'verified' if dim=='company' else 'unknown','reset_policy':'none','relation':'independent','status':'not_applicable' if dim in ('cooldown','change') else 'verified','text':'合成规则；无额外限制；公司上限单列','limit':len(roles)+5 if dim=='company' else '', 'used':0 if dim=='company' else '', 'role_ids':[], 'shared_group':'','source':'https://example.test/faq','checked_at':'20XX-10-16','gap':''} for dim in sorted(h.DIMENSIONS)],'referrals':[{'id':'ref','scope':'全部计划/BU','role_ids':[],'status':'not_found','code':'','source':'https://example.test/referral','checked_at':'20XX-10-16','expires':'','notes':'已查公开来源，无可核实码'}]}


class CompanyHandoffTests(unittest.TestCase):
 def setUp(self):
  self.tmp=tempfile.TemporaryDirectory();self.root=Path(self.tmp.name);self.book=self.root/'tracker.xlsx';self.batch='synthetic'
  wb=Workbook();ws=wb.active;ws.title='地区岗位';ws.append(['公司','岗位名称','所属BU','工作地点','岗位链接','岗位JD','投递状态']);ws.append(['合成公司','AI产品','BU-A','城市甲','https://example.test/job/1','职责：AI产品需求。要求：本科。','待投递']);ws.append(['合成公司','用户产品','BU-B','城市乙','https://example.test/job/2','职责：用户产品。要求：本科。','待投递']);wb.save(self.book)
  self.c=sample_company(self.book);self.c['roles'][1].update(tier='partial',selected=False,order='',recommended=True)
 def tearDown(self):self.tmp.cleanup()
 def draft(self):h.write(self.book,self.batch,{'companies':[self.c]})
 def approve(self):
  self.draft();(self.root/'confirmation.md').write_text('合成测试确认当前HTML的选择与全部规则；不是真实用户授权', encoding='utf-8')
  h.confirm(self.book,self.batch,self.c['company_key'],'confirmation.md',h.sha(h.paths(self.book,self.batch)/'公司核对.html'))
 def test_legacy_default_encoding_json_first_roundtrip(self):
  original_read,original_write=Path.read_text,Path.write_text
  def legacy_read(path,encoding=None,**kwargs):return original_read(path,encoding=encoding or 'cp1252',**kwargs)
  def legacy_write(path,data,encoding=None,**kwargs):return original_write(path,data,encoding=encoding or 'cp1252',**kwargs)
  with patch.object(Path,'read_text',legacy_read),patch.object(Path,'write_text',legacy_write):
   draft=self.root/'draft.json';draft.write_text(json.dumps({'companies':[self.c]},ensure_ascii=False),encoding='utf-8')
   with patch('sys.argv',['company_handoff.py','write','--workbook',str(self.book),'--batch',self.batch,'--input',str(draft)]):
    self.assertEqual(h.main(),0)
   (self.root/'confirmation.md').write_text('合成用户确认',encoding='utf-8')
   h.confirm(self.book,self.batch,self.c['company_key'],'confirmation.md',h.sha(h.paths(self.book,self.batch)/'公司核对.html'))
   self.assertEqual(h.check(self.book,self.batch,self.c['roles'][0]['url']),[])
 def test_full_flow_and_separate_selection(self):
  self.approve();self.assertEqual(h.check(self.book,self.batch,self.c['roles'][0]['url']),[])
  self.assertIn('岗位不在用户最终选择中',h.check(self.book,self.batch,self.c['roles'][1]['url']))
 def test_unconfirmed(self):
  self.draft();self.assertTrue(h.check(self.book,self.batch,self.c['roles'][0]['url']))
 def test_json_tamper(self):
  self.approve();p=h.paths(self.book,self.batch)/'公司方案.json';d=json.loads(p.read_text(encoding='utf-8'));d['companies'][0]['roles'][1]['selected']=True;p.write_text(json.dumps(d), encoding='utf-8');self.assertIn('JSON内容与Excel不一致',h.check(self.book,self.batch))
 def test_excel_refresh_stale_then_export(self):
  self.approve();wb=load_workbook(self.book);wb['地区岗位']['G2']='已提交';wb.save(self.book)
  self.assertIn('Excel/JSON未同步',h.check(self.book,self.batch));h.export(self.book,self.batch);self.assertEqual(h.check(self.book,self.batch),[])
 def test_business_change_invalidates_confirmation(self):
  self.approve();wb=load_workbook(self.book);ws=wb['投递规则'];ws['G2']='changed official rule';wb.save(self.book);h.export(self.book,self.batch)
  self.assertIn('公司方案未确认或确认已失效',h.check(self.book,self.batch))
 def test_html_tamper(self):
  self.approve();p=h.paths(self.book,self.batch)/'公司核对.html';p.write_text('changed', encoding='utf-8');self.assertIn('HTML与Excel不一致',h.check(self.book,self.batch))
 def test_confirmation_tamper(self):
  self.approve();(self.root/'confirmation.md').write_text('changed', encoding='utf-8');self.assertIn('用户确认文件变化',h.check(self.book,self.batch))
 def test_wrong_html_hash(self):
  self.draft();(self.root/'confirmation.md').write_text('synthetic', encoding='utf-8');
  with self.assertRaises(ValueError):h.confirm(self.book,self.batch,'synthetic','confirmation.md','wrong')
 def test_missing_rule_dimension(self):
  self.c['rules']=self.c['rules'][:-1]
  with self.assertRaises(ValueError):self.draft()
 def test_per_bu_scope_cannot_omit_other_candidates(self):
  next(x for x in self.c['rules'] if x['dimension']=='bu')['role_ids']=['r1']
  with self.assertRaises(ValueError):self.draft()
 def test_missing_referral(self):
  self.c['referrals']=[]
  with self.assertRaises(ValueError):self.draft()
 def test_per_scope_referral(self):
  self.c['referrals'][0]['role_ids']=['r1']
  with self.assertRaises(ValueError):self.draft()
 def test_special_pending(self):
  self.c['plans'][0].update(special=True,decision='pending')
  with self.assertRaises(ValueError):self.draft()
 def test_special_accepted(self):
  (self.root/'plan-confirm.md').write_text('synthetic: consider special plan', encoding='utf-8');self.c['plans'][0].update(special=True,decision='consider',decision_ref='plan-confirm.md');self.c['roles'][0]['tier']='special';self.approve();self.assertEqual(h.check(self.book,self.batch),[])
 def test_quota_each_dimension(self):
  for dim in ('company','round','bu','shared','plan','concurrent'):
   with self.subTest(dim=dim):
    c=copy.deepcopy(self.c);next(x for x in c['rules'] if x['dimension']==dim).update(limit=0,limit_type='count',role_ids=['r1'])
    with self.assertRaises(ValueError):h.write(self.book,self.batch,{'companies':[c]})
 def test_missing_regional_row(self):
  self.approve();wb=load_workbook(self.book);wb['地区岗位'].delete_rows(2);wb.save(self.book);h.export(self.book,self.batch);self.assertIn('地区岗位表缺少唯一对应行',h.check(self.book,self.batch))
 def test_regional_jd_changed(self):
  self.approve();wb=load_workbook(self.book);wb['地区岗位']['F2']='新要求';wb.save(self.book);h.export(self.book,self.batch);self.assertIn('地区表与确认目录不一致: 岗位JD',h.check(self.book,self.batch))
 def test_outside_directory_and_escaping(self):
  r=copy.deepcopy(self.c['roles'][1]);r.update(id='r3',title='<script>alert(1)</script>',tier='outside',jd='',selected=False,recommended=False,order='',url='',source='javascript:alert(1)');self.c['roles'].append(r);next(x for x in self.c['research'] if x['id']=='catalog').update(expected_count=3,observed_count=3);self.draft()
  page=(h.paths(self.book,self.batch)/'公司核对.html').read_text(encoding='utf-8');self.assertIn('&lt;script&gt;',page);self.assertNotIn('href="javascript:',page);self.assertIn('目录岗位 3 个',page)
 def test_preserve_other_sheets_and_formula_like_jd(self):
  self.c['roles'][1]['jd']='=HYPERLINK("x")';self.draft();wb=load_workbook(self.book);self.assertEqual(wb['地区岗位']['G2'].value,'待投递');self.assertEqual(wb['岗位目录']['L3'].data_type,'s')
 def test_long_text_rejected_not_truncated(self):
  self.c['roles'][0]['jd']='中'*32768
  with self.assertRaises(ValueError):self.draft()
 def test_same_version_refresh_rejected(self):
  self.draft()
  with self.assertRaises(ValueError):self.draft()
 def test_new_version_requires_new_confirmation(self):
  self.approve();self.c['version']='v2';self.draft();self.assertIn('公司方案未确认或确认已失效',h.check(self.book,self.batch))

 def test_empty_confirmation_rejected(self):
  self.draft();(self.root/'empty.md').write_text('', encoding='utf-8')
  with self.assertRaises(ValueError):h.confirm(self.book,self.batch,'synthetic','empty.md',h.sha(h.paths(self.book,self.batch)/'公司核对.html'))
 def test_duplicate_role_url_rejected(self):
  self.c['roles'][1]['url']=self.c['roles'][0]['url']
  with self.assertRaises(ValueError):self.draft()
 def test_expired_referral_blocks_handoff(self):
  self.c['referrals'][0].update(status='verified',cohort='测试届别',code='TEST',expires='2000-01-01');self.approve()
  self.assertIn('内推码已过期，须复核并同步',h.check(self.book,self.batch))
 def test_missing_json_after_interrupted_export(self):
  self.approve();(h.paths(self.book,self.batch)/'公司方案.json').unlink();self.assertTrue(h.check(self.book,self.batch));h.export(self.book,self.batch);self.assertEqual(h.check(self.book,self.batch),[])
 def test_refresh_keeps_previous_versions(self):
  self.approve();old=h.sha(h.paths(self.book,self.batch)/'公司方案.json');self.c['version']='v2';self.draft()
  self.assertTrue((h.paths(self.book,self.batch)/'历史'/(old+'.json')).is_file())

class FlowStorageTests(unittest.TestCase):
 setUp=CompanyHandoffTests.setUp
 tearDown=CompanyHandoffTests.tearDown
 draft=CompanyHandoffTests.draft
 approve=CompanyHandoffTests.approve
 def test_complete_snapshot_includes_unrelated_sheets_and_facts(self):
  import workbook_store as store
  from openpyxl.comments import Comment
  from datetime import date
  wb=load_workbook(self.book);ws=wb.create_sheet('保留表');ws.append(['自定义列','公式','日期']);ws.append(['原文','=1+2',date(2001,1,2)]);ws['A2'].comment=Comment('原始备注','synthetic');ws['A2'].hyperlink='https://example.test/source';ws.merge_cells('A4:B4');wb.save(self.book)
  self.draft();data=json.loads(store.full_path(self.book).read_text(encoding='utf-8'));extra=next(x for x in data['sheets'] if x['name']=='保留表');cells=extra['rows'][1]['cells']
  self.assertEqual(cells[0]['note']['text'],'原始备注');self.assertEqual(cells[1]['value'],'=1+2');self.assertIn('cached_value',cells[1]);self.assertEqual(cells[2]['value']['type'],'datetime');self.assertEqual(store.check_full(self.book),[])
 def test_unified_keyed_writer_and_readback(self):
  wb=load_workbook(self.book)
  for sheet,heads in [('校招主链接',['公司','地区','校招主链接']),('校招计划',['公司','地区','计划','依据链接'])]:wb.create_sheet(sheet).append(heads)
  wb.save(self.book)
  h.write(self.book,self.batch,{'companies':[self.c],'workbook_updates':[
   {'sheet':'地区岗位','key':{'公司':'合成公司','岗位链接':self.c['roles'][0]['url']},'values':{'岗位JD':self.c['roles'][0]['jd']}},
   {'sheet':'校招主链接','key':{'公司':'合成公司','地区':'合成地区'},'values':{'校招主链接':'https://example.test/campus'}},
   {'sheet':'校招计划','key':{'公司':'合成公司','地区':'合成地区','计划':'普通校招'},'values':{'依据链接':'https://example.test/faq'}}]})
  import workbook_store as store
  rows=store.query(self.book,'合成公司');self.assertIn('校招主链接',{x['sheet'] for x in rows});self.assertIn('研究检查',{x['sheet'] for x in rows});self.assertEqual(store.check_full(self.book),[])
 def test_typed_date_writes_real_excel_date(self):
  wb=load_workbook(self.book);wb['地区岗位']['H1']='核验日期';wb.save(self.book)
  h.write(self.book,self.batch,{'companies':[self.c],'workbook_updates':[{'sheet':'地区岗位','key':{'岗位链接':self.c['roles'][0]['url']},'values':{'核验日期':{'type':'date','value':'2001-01-02'}}}]})
  wb=load_workbook(self.book);self.assertEqual(wb['地区岗位']['H2'].value.date().isoformat(),'2001-01-02');wb.close()
 def test_failed_upsert_does_not_commit_partial_workbook(self):
  before=h.sha(self.book)
  with self.assertRaises(ValueError):h.write(self.book,self.batch,{'companies':[self.c],'workbook_updates':[{'sheet':'地区岗位','key':{'公司':'合成公司','岗位链接':self.c['roles'][0]['url']},'values':{'岗位名称':'改变'}},{'sheet':'地区岗位','key':{'公司':'合成公司'},'values':{'不存在':'x'}}]})
  self.assertEqual(h.sha(self.book),before)
 def test_small_patch_preserves_other_role_data(self):
  self.draft();h.write(self.book,self.batch,{'company_patches':[{'company_key':'synthetic','version':'v2','roles':[{'id':'r2','recommended':False}]}]})
  c=h.read(self.book,self.batch)[0];self.assertFalse(c['roles'][1]['recommended']);self.assertEqual(c['roles'][0]['jd'],self.c['roles'][0]['jd'])
 def test_string_false_is_not_silently_selected(self):
  self.c['roles'][1]['selected']='false'
  with self.assertRaises(ValueError):self.draft()
 def test_unknown_payload_not_silently_discarded(self):
  self.c['roles'][0]['volunteer']='一志愿'
  with self.assertRaises(ValueError):self.draft()
 def test_complete_json_tampering_blocks(self):
  import workbook_store as store
  self.approve();path=store.full_path(self.book);data=json.loads(path.read_text(encoding='utf-8'));data['sheets'][0]['rows'][1]['cells'][0]['value']='篡改';path.write_text(json.dumps(data),encoding='utf-8')
  self.assertTrue(h.check(self.book,self.batch))
  with self.assertRaises(ValueError):store.query(self.book,'合成公司')
 def test_refresh_updates_all_batches(self):
  self.draft();h.write(self.book,'second',{'companies':[self.c]});h.export(self.book,self.batch)
  for batch in (self.batch,'second'):self.assertEqual(json.loads((h.paths(self.book,batch)/'公司方案.json').read_text(encoding='utf-8'))['workbook_sha256'],h.sha(self.book))
 def test_legacy_empty_string_is_preserved(self):
  import workbook_store as store
  wb=load_workbook(self.book);wb['地区岗位']['H1']='保留列';wb['地区岗位']['H2']='';wb.save(self.book)
  store.preserve_empty_text(self.book,{'地区岗位':{'H2'}})
  before=store.snapshot(self.book)['sheets'];self.draft();after=store.snapshot(self.book)['sheets']
  self.assertEqual(before,after[:len(before)])
 def test_empty_preference_is_explicit_null(self):
  self.approve();c=h.read(self.book,self.batch)[0];self.assertIsNone(c['preferences'][0]['rank']);self.assertEqual(h.check(self.book,self.batch,c['roles'][0]['url'],target='无'),[])
 def test_unknown_preference_blocks_handoff(self):
  self.c['preferences'][0]['status']='unknown';self.draft();self.assertIn('已选岗位志愿待核实',h.check(self.book,self.batch))
 def test_preference_not_processing_order(self):
  for idx,v in enumerate(self.c['preferences']):v.update(status='selected',mode='ordered',group_id='regular',rank=2-idx,label='产品志愿',bu='BU',cities='城市')
  next(r for r in self.c['rules'] if r['dimension']=='preference')['preference_mode']='ordered'
  self.approve();url=self.c['roles'][0]['url'];self.assertEqual(h.check(self.book,self.batch,url,target='regular/2/产品志愿/BU/城市'),[]);self.assertIn('目标确认的志愿与公司方案不一致',h.check(self.book,self.batch,url,target='regular/1/产品志愿/BU/城市'))
 def test_preference_duplicate_within_group_rejected(self):
  for v in self.c['preferences']:v.update(status='selected',mode='ordered',group_id='regular',rank=1,label='志愿')
  next(r for r in self.c['rules'] if r['dimension']=='preference')['preference_mode']='ordered'
  with self.assertRaises(ValueError):self.draft()
 def test_parallel_preferences_keep_null_rank(self):
  for v in self.c['preferences']:v.update(status='selected',mode='parallel',group_id='campus',rank=None,label='平行志愿',bu='BU',cities='城市')
  next(r for r in self.c['rules'] if r['dimension']=='preference')['preference_mode']='parallel'
  self.approve();url=self.c['roles'][0]['url'];self.assertEqual(h.check(self.book,self.batch,url,target='campus/parallel/平行志愿/BU/城市'),[])
 def test_same_preference_rank_in_different_groups_allowed(self):
  for idx,v in enumerate(self.c['preferences']):v.update(status='selected',mode='ordered',group_id='plan-'+str(idx),rank=1,label='志愿')
  next(r for r in self.c['rules'] if r['dimension']=='preference')['preference_mode']='ordered'
  self.approve();self.assertEqual(h.check(self.book,self.batch),[])
 def test_research_incomplete_cannot_claim_complete(self):
  self.c['research'][0].update(status='todo',gap='还未搜索')
  with self.assertRaises(ValueError):self.draft()
  self.c.update(coverage='partial',gaps='还未搜索');self.draft();self.assertTrue(any('研究未完成' in x for x in h.check(self.book,self.batch)))
 def test_all_structured_dimensions_roundtrip(self):
  import workbook_store as store
  cool=next(x for x in self.c['rules'] if x['dimension']=='cooldown')
  cool.update(status='verified',cooldown_triggers=['rejected'],cooldown_duration=90,cooldown_unit='days',cooldown_release='期满且账号显示可重新申请')
  self.approve();c=h.read(self.book,self.batch)[0];saved=next(x for x in c['rules'] if x['dimension']=='cooldown')
  self.assertEqual(saved['cooldown_triggers'],['rejected']);self.assertEqual(saved['cooldown_duration'],90)
  full=json.loads(store.full_path(self.book).read_text(encoding='utf-8'));self.assertEqual(full['company_batches'][self.batch],[c])
 def test_other_requirements_are_typed_atomic_conditions(self):
  self.c['conditions']=[{'id':'training-years','subject_type':'plan','subject_id':'regular','field':'training_duration','operator':'eq','condition_value':2,'unit':'years','group_id':'training','group_operator':'all','importance':'informational','status':'verified','source':'https://example.test/plan','checked_at':'20XX-10-16','gap':'','raw':'培养期两年'}]
  self.approve();c=h.read(self.book,self.batch)[0];self.assertEqual(c['conditions'][0]['condition_value'],2)
 def test_free_text_does_not_substitute_structured_rule(self):
  cool=next(x for x in self.c['rules'] if x['dimension']=='cooldown');cool.update(status='verified',text='被拒后九十天才能重投')
  with self.assertRaises(ValueError):self.draft()
 def test_structured_plan_bu_rule_scope_is_enforced(self):
  rule=next(x for x in self.c['rules'] if x['dimension']=='bu');rule['bu_ids']=['BU-A']
  with self.assertRaises(ValueError):self.draft()
 def test_catalog_count_mismatch(self):
  next(x for x in self.c['research'] if x['id']=='catalog')['expected_count']=50
  with self.assertRaises(ValueError):self.draft()
 def test_init_preserves_existing_facts_and_no_consent(self):
  import workbook_store as store
  before=store.snapshot(self.book)['sheets'];result=store.initialize(self.book,self.batch,'地区岗位','已确认合成范围')
  self.assertEqual(result['initialized_or_migrated'],1);after=store.snapshot(self.book)['sheets'];self.assertEqual(before,after[:len(before)])
  c=h.read(self.book,self.batch)[0];self.assertEqual(c['roles'],[]);self.assertEqual(c['confirmed_at'],'');self.assertEqual(len(c['existing_data']),2);self.assertEqual(sum(x.startswith('研究未完成:') for x in h.audit(self.book,self.batch)[0]['issues']),7)
  self.assertEqual(store.initialize(self.book,self.batch,'地区岗位','已确认合成范围')['initialized_or_migrated'],0)
 def test_v1_headers_migrate_to_structured_dimensions(self):
  import workbook_store as store
  from structured_dimensions import EXTRA_FIELDS
  self.approve();wb=load_workbook(self.book)
  for kind,fields in EXTRA_FIELDS.items():
   sheet=wb[h.SCHEMA[kind][0]];original=len(h.headers(kind))-len(fields);sheet.delete_cols(original+1,len(fields))
  del wb['志愿明细'];del wb['研究检查'];wb.save(self.book)
  store.initialize(self.book,self.batch,'地区岗位','合成范围');c=h.read(self.book,self.batch)[0]
  self.assertEqual(c['scope_status'],'unknown');self.assertEqual(c['roles'][0]['jd'],self.c['roles'][0]['jd']);self.assertEqual(c['confirmed_at'],'')
  self.assertTrue(h.check(self.book,self.batch))
 def test_old_schema_migrates_without_preserving_approval(self):
  import workbook_store as store
  self.approve();wb=load_workbook(self.book);del wb['志愿明细'];del wb['研究检查'];wb.save(self.book)
  store.initialize(self.book,self.batch,'地区岗位','合成范围');c=h.read(self.book,self.batch)[0];self.assertEqual(c['confirmed_at'],'');self.assertEqual(c['preferences'][0]['status'],'unknown');self.assertTrue(h.check(self.book,self.batch))
 def test_submitted_status_never_reset_by_writer(self):
  wb=load_workbook(self.book);wb['地区岗位']['G2']='已提交';wb.save(self.book);before=h.sha(self.book)
  with self.assertRaises(ValueError):h.write(self.book,self.batch,{'companies':[self.c],'workbook_updates':[{'sheet':'地区岗位','key':{'岗位链接':self.c['roles'][0]['url']},'values':{'投递状态':'待投递'}}]})
  self.assertEqual(h.sha(self.book),before)

if __name__=='__main__':unittest.main()
