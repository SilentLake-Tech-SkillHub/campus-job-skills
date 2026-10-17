"""Deterministic company funnel summary shared by phase and workbook HTML views."""
from html import escape


def _unique(records):
    result = {}
    for record in records:
        key = record.get('id') or record.get('url')
        if not key:
            raise ValueError('汇报条目缺少官方ID或真实链接')
        if key in result and result[key] != record:
            raise ValueError('同一身份存在冲突记录，不能汇总')
        result[key] = record
    return list(result.values())


def summarize(company):
    plans = _unique(company.get('plans', []))
    roles = _unique(company.get('roles', []))
    checks = {x['id']: x for x in company.get('research', [])}
    def checked(key):
        x = checks.get(key, {})
        return (x.get('status') == 'checked' and not x.get('gap')
                and (key not in ('plans', 'catalog') or
                     type(x.get('expected_count')) is int and
                     x['expected_count'] == x.get('observed_count')))
    plans_done = checked('plans')
    roles_done = plans_done and checked('catalog') and checked('jd')
    candidates = [r for r in roles if r.get('tier') in ('exact', 'partial')
                  or r.get('tier') == 'special' and r.get('functional_fit') in ('exact', 'partial', '符合', '部分符合')]
    trainee = [p for p in plans if p.get('category') in ('management_trainee', 'product_trainee')]
    trainee_done = plans_done and all(p.get('category') in
        ('regular', 'technical', 'product_trainee', 'management_trainee', 'other') for p in plans)
    conditions = company.get('conditions', [])
    def conversion(r):
        return any(c.get('subject_type') == 'role' and c.get('subject_id') == r['id']
                   and c.get('field') == 'conversion_opportunity'
                   and c.get('condition_value') is True and c.get('status') == 'verified'
                   and c.get('source') and c.get('raw') for c in conditions)
    interns = [r for r in candidates if r.get('employment') in ('实习', 'internship')]
    conversions = [r for r in interns if conversion(r)]
    # A missing conversion fact cannot silently turn into a verified zero.
    conversion_done = roles_done and all(conversion(r) or any(
        c.get('subject_id') == r['id'] and c.get('subject_type') == 'role'
        and c.get('field') == 'conversion_opportunity' and c.get('condition_value') is False
        and c.get('status') == 'verified' and c.get('source') and c.get('raw')
        for c in conditions) for r in interns)
    def count(items, done):
        return {'count': len(items) if done else None, 'observed_count': len(items),
                'status': 'verified' if done else 'pending'}
    return {'company': company['name'], 'scope': company.get('scope', ''),
            'plans': {**count(plans, plans_done), 'names': [p['name'] for p in plans]},
            'matching_roles': count(candidates, roles_done),
            'formal_roles': count([r for r in candidates if r.get('employment') in ('正式', 'full_time')], roles_done),
            'trainee_plans': {**count(trainee, trainee_done), 'names': [p['name'] for p in trainee]},
            'conversion_roles': count(conversions, conversion_done),
            'selected_roles': sum(r.get('selected') is True for r in roles)}


def company_scan_complete(company):
    """Only the whole-company research checkpoint can publish this report."""
    checks = {x['id']: x for x in company.get('research', [])}
    for key in ('plans', 'catalog', 'jd', 'eligibility', 'rules', 'referrals', 'preferences'):
        x = checks.get(key, {})
        if x.get('status') != 'checked' or x.get('gap'):
            return False
        if key in ('plans', 'catalog') and (type(x.get('expected_count')) is not int
                or x['expected_count'] != x.get('observed_count')):
            return False
    return True


def render_summary(company):
    if not company_scan_complete(company):
        return ''
    s = summarize(company)
    def number(item):
        return str(item['count']) if item['status'] == 'verified' else '待核（已发现 '+str(item['observed_count'])+'）'
    names = '、'.join(s['plans']['names']) or ('无' if s['plans']['status'] == 'verified' else '待核')
    trainees = '、'.join(s['trainee_plans']['names'])
    esc = lambda v: escape(str(v))
    return ('<div data-company-summary="true"><h3>公司范围与岗位汇报</h3>'
            '<p>当前公司：'+esc(s['company'])+'</p><p>校招范围：'+esc(names)+'；招聘计划 '+esc(number(s['plans']))+' 个。</p>'
            '<p>符合您岗位方向及硬范围要求的岗位：'+esc(number(s['matching_roles']))+' 个，其中正式岗 '+esc(number(s['formal_roles']))+' 个。</p>'
            '<p>管培生／产品培训生计划：'+esc(number(s['trainee_plans']))+' 个'+('，分别是：'+esc(trainees) if trainees else '')+'。</p>'
            '<p>符合范围的实习转正岗位：'+esc(number(s['conversion_roles']))+' 个。</p>'
            '<p><small>统计范围：'+esc(s['scope'])+'。实习转正岗位包含在候选总数中，不重复相加；以上是岗位方向统计，个人资格另核。用户选择投递 '+str(s['selected_roles'])+' 个岗位，独立记录。核完无结果显示0，未核完显示待核；入职培训不算招聘计划。</small></p></div>')
