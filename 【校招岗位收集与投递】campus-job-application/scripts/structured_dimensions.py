"""Typed dimensions shared by workbook, company JSON, review and application gates."""
EXTRA_FIELDS = {
 'company': [('scope_status','需求结构化状态'),('regions','地区范围'),('cohort','招聘届别'),('employment_types','用工范围'),('role_families','岗位族'),('hard_cities','城市硬范围'),('preferred_cities','城市偏好'),('priority_scope','用户志愿排序范围'),('priority_scope_source','排序范围来源'),('priority_scope_checked_at','排序范围核验时间')],
 'plans': [('category','计划类别'),('window_status','窗口状态'),('window_start','开放日期'),('window_end','截止日期'),('requirements_status','要求结构化状态'),('degree_levels','学历要求'),('majors','专业要求'),('graduation_start','毕业起始日期'),('graduation_end','毕业截止日期')],
 'roles': [('round_id','招聘轮次ID'),('city_options','城市选项')],
 'rules': [('preference_mode','志愿机制'),('plan_ids','适用计划ID'),('bu_ids','适用BU'),('round_id','适用轮次ID'),('round_start','轮次开始日期'),('round_end','轮次结束日期'),('limit_type','限额类型'),('usage_status','已用名额状态'),('reset_policy','名额重置机制'),('reset_at','重置日期'),('relation','计划或BU关系'),('cooldown_triggers','冷冻触发动作'),('cooldown_duration','冷冻时长'),('cooldown_unit','冷冻时长单位'),('cooldown_until','冷冻截止日期'),('cooldown_release','解除条件'),('change_actions','撤回换岗动作'),('quota_effect','动作对名额影响')],
 'referrals': [('plan_ids','适用计划ID'),('bu_ids','适用BU'),('cohort','适用届别'),('valid_from','生效日期')],
}
LIST_FIELDS={'regions','employment_types','role_families','hard_cities','preferred_cities','degree_levels','majors','city_options','bu_ids','cooldown_triggers','change_actions'}
NULL_FIELDS={'rank','cooldown_duration'}
ENUMS={'reset_policy':{'none','round','annual','date','unknown','not_applicable'},'relation':{'independent','shared','exclusive','conditional','unknown','not_applicable'},'usage_status':{'verified','unknown','not_applicable'},'quota_effect':{'retained','released','consumed','conditional','unknown','not_applicable'}}
DATE_FIELDS={'window_start','window_end','graduation_start','graduation_end','round_start','round_end','reset_at','cooldown_until','valid_from'}
QUOTA_DIMENSIONS={'company','plan','round','bu','shared','concurrent','preference'}


def applies(record,role):
 return ((not record.get('role_ids') or role['id'] in record['role_ids'])
  and (not record.get('plan_ids') or bool(set(role['plan_ids']) & set(record['plan_ids'])))
  and (not record.get('bu_ids') or role['bu'] in record['bu_ids'])
  and (not record.get('round_id') or role.get('round_id')==record['round_id']))


def validate(c,ready=False):
 errors=[];complete=ready or c.get('coverage')=='complete';plans={p['id'] for p in c['plans']}
 if c.get('scope_status') not in ('verified','unknown'):errors.append('需求必须结构化或明确待补')
 if complete and (c.get('scope_status')!='verified' or not all(c.get(k) for k in ('regions','cohort','employment_types','role_families'))):errors.append('交接缺结构化需求范围')
 from datetime import date
 for kind in ('company','plans','roles','rules','referrals'):
  for item in ([c] if kind=='company' else c[kind]):
   for key in DATE_FIELDS:
    if item.get(key):
     try:date.fromisoformat(item[key])
     except (ValueError,TypeError):errors.append('日期维度须为ISO日期: '+key)
   for key,allowed in ENUMS.items():
    if item.get(key) and item[key] not in allowed:errors.append('枚举维度无效: '+key)
   for key in LIST_FIELDS|{'plan_ids','role_ids'}:
    if key in item and not isinstance(item[key],list):errors.append(kind+' 列表维度类型错误: '+key)
 for p in c['plans']:
  if complete:
   if p.get('category') not in ('regular','technical','product_trainee','management_trainee','other'):errors.append('计划类别尚未结构化')
   if p.get('window_status') not in ('known','not_disclosed'):errors.append('计划窗口尚未结构化')
   if p.get('window_status')=='known' and not (p.get('window_start') or p.get('window_end')):errors.append('已知窗口缺日期')
   if p.get('requirements_status') not in ('structured','not_disclosed'):errors.append('计划要求尚未结构化')
 for r in c['roles']:
  if complete and r['tier'] in ('exact','partial','special') and not r.get('round_id'):errors.append('候选缺招聘轮次ID或明确未披露值')
 for r in c['rules']:
  dim=r['dimension'];status=r['status']
  if not set(r.get('plan_ids',[]))<=plans:errors.append('规则引用未知计划')
  if dim in QUOTA_DIMENSIONS and (complete or status=='verified'):
   kind=r.get('limit_type')
   if kind not in ('count','unlimited','unknown','not_applicable'):errors.append('限额类型尚未结构化')
   if kind=='count' and (type(r.get('limit'))!=int or r['limit']<0):errors.append('数量限额缺非负整数')
   if kind!='count' and r.get('limit') not in ('',None):errors.append('非数量限额不能填数量')
   if kind=='unknown' and status!='unknown':errors.append('未知限额不可标已核实')
  if r.get('usage_status')=='verified' and (type(r.get('used'))!=int or r['used']<0):errors.append('已核实名额缺非负整数')
  if r.get('used') not in ('',None) and r.get('usage_status')!='verified':errors.append('已用名额缺核验状态')
  if complete and dim=='preference' and status=='verified' and r.get('preference_mode') not in ('ordered','parallel','not_applicable'):errors.append('志愿规则机制尚未结构化')
  if complete and dim=='round' and status=='verified' and (not r.get('round_id') or not r.get('reset_policy')):errors.append('轮次及名额重置尚未结构化')
  if complete and dim in ('plan','shared') and status=='verified' and not r.get('relation'):errors.append('跨计划关系尚未结构化')
  if complete and dim=='cooldown' and status=='verified':
   if not r.get('cooldown_triggers') or r.get('cooldown_unit') not in ('days','months','years','cycle','until_date','conditional'):errors.append('冷冻触发和时长单位尚未结构化')
   if r.get('cooldown_unit') in ('days','months','years') and (type(r.get('cooldown_duration'))!=int or r['cooldown_duration']<0):errors.append('冷冻时长缺非负整数')
   if r.get('cooldown_unit')=='until_date' and not r.get('cooldown_until'):errors.append('冷冻截止日期缺失')
   if not r.get('cooldown_release'):errors.append('冷冻解除条件缺失')
  if complete and dim=='change' and status=='verified' and (not r.get('change_actions') or not r.get('quota_effect')):errors.append('撤回换岗后果尚未结构化')
 for r in c['referrals']:
  if not set(r.get('plan_ids',[]))<=plans:errors.append('内推引用未知计划')
  if r['status']=='verified' and r.get('cohort')!=c.get('cohort'):errors.append('内推适用届别不一致或缺失')
 subjects={'company':{c['company_key']},'plan':plans,'role':{r['id'] for r in c['roles']},'rule':{r['id'] for r in c['rules']}}
 for fact in c.get('conditions',[]):
  if fact.get('subject_id') not in subjects.get(fact.get('subject_type'),set()):errors.append('条件引用未知对象')
  if not fact.get('field') or fact['field'] in ('other','其他','备注'):errors.append('条件必须命名具体维度')
  if fact.get('operator') not in ('eq','neq','in','not_in','gte','lte','between','exists'):errors.append('条件比较关系无效')
  if fact.get('group_operator') not in ('all','any') or not fact.get('group_id'):errors.append('条件组缺少与或关系')
  if fact.get('importance') not in ('required','preferred','informational'):errors.append('条件必须区分必需、优先、说明')
  if fact.get('status') not in ('verified','unknown','not_applicable'):errors.append('条件状态无效')
  value=fact.get('condition_value');op=fact.get('operator')
  if fact.get('status')=='verified':
   if value is None or not fact.get('source') or not fact.get('checked_at'):errors.append('已核实条件缺值或来源时间')
   if op in ('gte','lte') and (type(value) not in (int,float)):errors.append('数值条件必须存数值')
   if op in ('in','not_in','between') and not isinstance(value,list):errors.append('集合或区间条件必须存数组')
   if op=='between' and isinstance(value,list) and len(value)!=2:errors.append('区间条件须有两个端点')
  if fact.get('status')=='unknown' and not fact.get('gap'):errors.append('未知条件缺待补事实')
 return errors
