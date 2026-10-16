"""User-confirmed company/BU ordering, independent of portal preference slots."""
FIELDS = [('id','用户志愿ID'),('role_id','岗位ID'),('scope','排序范围'),
          ('group_id','用户志愿组'),('bu','排序BU'),('rank','用户志愿序号'),
          ('status','用户确认状态'),('confirmation_ref','用户确认依据'),
          ('confirmation_sha256','用户确认依据哈希'),('confirmed_at','用户确认时间')]


def validate(c, ready=False):
    errors=[]
    scope=c.get('priority_scope') or 'unknown'
    if scope not in ('company','bu','unknown'):
        errors.append('用户志愿排序范围必须为 company/bu/unknown')
    roles={r['id']:r for r in c['roles']}
    seen=set();slots=set();groups={}
    for p in c.get('user_priorities',[]):
        rid=p.get('role_id');rank=p.get('rank');status=p.get('status')
        if rid not in roles or rid in seen:
            errors.append('用户志愿引用未知或重复岗位')
        seen.add(rid)
        if p.get('scope') not in ('company','bu','unknown') or p.get('scope')!=scope:
            errors.append('用户志愿分组与公司排序范围不一致')
        if status not in ('pending','confirmed','unknown'):
            errors.append('用户志愿确认状态无效')
        bu=p.get('bu','')
        expected=c['company_key']+('::'+bu if scope=='bu' else '')
        if scope=='company' and bu:
            errors.append('公司内志愿组不能分BU')
        if scope=='bu' and (not bu or bu in ('未披露','unknown') or (rid in roles and bu!=roles[rid]['bu'])):
            errors.append('BU内志愿必须对应岗位的已核实BU')
        if scope!='unknown' and p.get('group_id')!=expected:
            errors.append('用户志愿组身份不匹配')
        if rank is not None:
            if type(rank)!=int or rank<1:
                errors.append('用户志愿序号须为正整数或null')
            else:
                slot=(p.get('group_id'),rank)
                if slot in slots:errors.append('同一用户志愿组序号重复')
                slots.add(slot)
                groups.setdefault(p.get('group_id'),[]).append(rank)
        if status=='confirmed':
            if scope=='unknown' or rank is None or not all(p.get(k) for k in ('confirmation_ref','confirmation_sha256','confirmed_at')):
                errors.append('已确认用户志愿缺分组、序号或确认依据')
        elif rank is not None:
            errors.append('未确认用户志愿不能填确定序号')
    for ranks in groups.values():
        if sorted(ranks)!=list(range(1,len(ranks)+1)):
            errors.append('用户志愿组序号须从1连续排列')
    if ready:
        for r in c['roles']:
            if r['selected']:
                if scope=='unknown':errors.append('已选岗位的公司或BU排序范围待核')
                if not any(p.get('role_id')==r['id'] and p.get('status')=='confirmed' for p in c.get('user_priorities',[])):
                    errors.append('已选岗位缺用户确认的公司或BU志愿顺序: '+r['id'])
        if scope!='unknown' and (not c.get('priority_scope_source') or not c.get('priority_scope_checked_at')):
            errors.append('用户志愿分组范围缺核验来源或日期')
    return errors


def evidence_errors(c,root,safe_ref,sha):
    errors=[]
    for p in c.get('user_priorities',[]):
        if p.get('status')!='confirmed':continue
        try:
            if sha(safe_ref(root,p['confirmation_ref']))!=p['confirmation_sha256']:
                errors.append('用户志愿确认依据已变化')
        except (OSError,ValueError,KeyError):
            errors.append('用户志愿确认依据不可读取')
    return errors
