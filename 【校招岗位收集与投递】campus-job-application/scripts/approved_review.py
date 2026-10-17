"""Bind one approved HTML to structured data; refreshing JSON never rewrites it."""
from pathlib import Path

FIELDS=[('review_mode','核对页模式'),('review_html_ref','最终HTML路径'),('review_html_sha256','最终HTML哈希'),('review_evidence_ref','方案确认依据'),('review_evidence_sha256','方案确认依据哈希'),('reviewed_at','方案确认记录时间'),('review_business_sha256','确认方案数据摘要')]
ROLE_FIELDS=[('application_decision','逐岗投递决定'),('decision_ref','逐岗决定依据'),('decision_sha256','逐岗决定依据哈希')]
REVIEW_FIELDS={x[0] for x in FIELDS}

def business_digest(c):
    from company_handoff import CONFIRM_FIELDS, SCHEMA, digest
    # Canonicalize absent optional columns exactly like workbook readback.
    result={}
    for kind,(_,fields) in SCHEMA.items():
        if kind=='company':
            result.update({key:c.get(key,'') for key,_ in fields if key not in CONFIRM_FIELDS|REVIEW_FIELDS})
        else:result[kind]=c.get(kind,[])
    return digest(result)

def validate_binding(c,root):
    from company_handoff import safe_ref, sha
    if c.get('review_mode')!='frozen':return []
    errors=[]
    for ref,expected,label in [('review_html_ref','review_html_sha256','已确认HTML'),('review_evidence_ref','review_evidence_sha256','用户确认依据')]:
        try:
            if sha(safe_ref(Path(root),c.get(ref)))!=c.get(expected):errors.append(label+'哈希不一致')
        except (OSError,ValueError):errors.append(label+'不可读取')
    for role in c.get('roles',[]):
        if role.get('application_decision') not in ('apply','decline','defer','second_intent'):continue
        try:
            if sha(safe_ref(Path(root),role.get('decision_ref')))!=role.get('decision_sha256'):errors.append('逐岗决定依据哈希不一致')
        except (OSError,ValueError):errors.append('逐岗决定依据不可读取')
    if not c.get('reviewed_at'):errors.append('缺少用户方案确认记录时间')
    if c.get('review_business_sha256')!=business_digest(c):errors.append('数据与已确认HTML绑定的方案不一致')
    return errors

def validate_decisions(c):
    errors=[]
    for r in c.get('roles',[]):
        decision=r.get('application_decision','')
        if not decision:continue  # Legacy data retains its original meaning.
        if decision not in ('apply','decline','defer','second_intent','pending'):errors.append('逐岗投递决定无效')
        if decision=='apply' and not r.get('selected'):errors.append('投递决定与用户选择不一致')
        if decision in ('decline','defer','second_intent','pending') and r.get('selected'):errors.append('未选择投递的岗位不可设为已选')
        if decision in ('apply','decline','defer','second_intent') and not (r.get('decision_ref') and r.get('decision_sha256')):errors.append('逐岗决定缺人工依据')
    return errors

def manifest(c):
    return {key:c.get(key,'') for key in REVIEW_FIELDS}
