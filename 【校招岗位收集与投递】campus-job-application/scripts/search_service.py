#!/usr/bin/env python3
"""campus-search MCP：多浏览器搜索任务队列，以及主数据 JSON 的固定读写接口（薄封装 json_excel_sync.py）。"""
import json
import os
import sys
from search_queue import Queue
from search_paths import configure


OPERATIONS={}
def operation(fn):
    OPERATIONS[fn.__name__]=fn
    return fn

def execute(name,arguments):
    configure()
    import inspect
    fn=OPERATIONS.get(name)
    if fn is None:raise ValueError('UNKNOWN_OPERATION')
    inspect.signature(fn).bind(**arguments)
    return fn(**arguments)

SYNC_SCRIPT=os.path.join(os.path.dirname(os.path.abspath(__file__)),'json_excel_sync.py')


def run_requirements(script, args):
    import subprocess
    workspace=os.environ.get('CAMPUS_RECRUITING_WORKSPACE')
    if not workspace:raise ValueError('需要配置CAMPUS_RECRUITING_WORKSPACE')
    cmd=[sys.executable,'-B',os.path.join(os.path.dirname(os.path.abspath(__file__)),script),
         '--workspace',workspace,*args]
    p=subprocess.run(cmd,capture_output=True,text=True,encoding='utf-8')
    if p.returncode:raise ValueError((p.stderr or p.stdout).strip())
    return json.loads(p.stdout)


@operation
def recruiting_application_requirements_begin(session_id:str|None=None,turn_id:str|None=None)->dict:
    """搜索组件启动：读取本回合需求入口；同回合默认模式不重复询问，新回合重新初始化。"""
    args=[]
    if session_id:args+=['--session',session_id]
    if turn_id:args+=['--turn',turn_id]
    return run_requirements('application_requirements.py',args+['begin'])


@operation
def recruiting_application_requirements_set(requirements:dict,user_quote:str,session_id:str|None=None,turn_id:str|None=None)->dict:
    """覆盖写入本轮需求临时JSON（公司/地区/浏览器）；空对象表示用户选择默认，不改主表状态。"""
    args=[]
    if session_id:args+=['--session',session_id]
    if turn_id:args+=['--turn',turn_id]
    return run_requirements('application_requirements.py',args+['set','--requirements',json.dumps(requirements,ensure_ascii=False),'--quote',user_quote])


@operation
def recruiting_application_forward(session_id:str|None=None,turn_id:str|None=None)->dict:
    """读取一次性需求生成当前正向评审；无公司参数默认取最新十家。成功后清空需求表，失败保留，仍须用户确认正向评审。"""
    args=[]
    if session_id:args+=['--session',session_id]
    if turn_id:args+=['--turn',turn_id]
    return run_requirements('company_store.py',args+['forward'])


def run_sync(*args, items=None):
    """调用唯一的主数据读写脚本；所有门禁在脚本内执行。"""
    import subprocess, tempfile
    workspace=os.environ.get('CAMPUS_RECRUITING_WORKSPACE')
    if not workspace:raise ValueError('需要配置CAMPUS_RECRUITING_WORKSPACE')
    cmd=[sys.executable,'-B',SYNC_SCRIPT,'--workspace',workspace,*args]
    tmp=None
    if items is not None:
        tmp=tempfile.NamedTemporaryFile('w',suffix='.json',delete=False,encoding='utf-8');json.dump(items,tmp,ensure_ascii=False);tmp.close();cmd+=['--input',tmp.name]
    try:
        p=subprocess.run(cmd,capture_output=True,text=True,encoding='utf-8')
    finally:
        if tmp:os.unlink(tmp.name)
    if p.returncode not in (0,1) or (p.returncode==1 and not p.stdout.strip()):raise ValueError((p.stderr or p.stdout).strip())
    return json.loads(p.stdout)


@operation
def recruiting_status_propose(items:list[dict])->dict:
    """只读：生成“当前状态是X，更新状态是Y”（含名额收口派生变更）。items：[{code: 岗位编码, to}]，公司行 [{sheet, code: 公司编码, to}]。必须原样展示给用户并取得明确确认。"""
    return run_sync('propose',items=items)


@operation
def recruiting_status_confirm(items:list[dict],user_quote:str,confirmation_ref:str='')->dict:
    """用户明确确认 propose 展示的变更后调用：记入确认台账并更新主数据 JSON，不写 Excel。"""
    return run_sync('confirm','--quote',user_quote,'--ref',confirmation_ref,items=items)


@operation
def recruiting_rows_upsert(items:list[dict])->dict:
    """新增岗位（{key:{sheet,company,role:"new"},values}，脚本分配岗位编码）或按编码改非状态字段（{key:{code},values}）；公司行用 {key:{sheet,company}}。
    不能改投递状态、岗位链接（recruiting_link_change_*）、投递链接、站点链接、志愿顺序和编码列。"""
    return run_sync('upsert',items=items)


@operation
def recruiting_link_exception_preview(items:list[dict])->dict:
    """只读：登记或撤销链接例外（官网无独立详情页）前生成备注“当前 → 更新”（items：[{code: 岗位编码, 定位说明}]，撤销给 "撤销": true）。必须原样展示给用户并取得明确确认。"""
    return run_sync('link-exception','--preview',items=items)


@operation
def recruiting_link_exception_confirm(items:list[dict],user_quote:str)->dict:
    """用户确认后写入固定格式的链接例外备注并记入台账；forward 据此不检查该岗位链接是否为具体岗位页。不写 Excel。"""
    return run_sync('link-exception','--quote',user_quote,items=items)


@operation
def recruiting_link_change_preview(items:list[dict])->dict:
    """只读：更换已有岗位链接前生成“岗位链接 X → Y”（items：[{code: 岗位编码, new_link}]）。必须原样展示给用户并取得明确确认。"""
    return run_sync('relink','--preview',items=items)


@operation
def recruiting_link_change_confirm(items:list[dict],user_quote:str)->dict:
    """用户确认后更换已有岗位链接：岗位编码与投递状态不变，引用它的记录都不受影响；不写 Excel。"""
    return run_sync('relink','--quote',user_quote,items=items)


@operation
def recruiting_apply_link_preview(items:list[dict])->dict:
    """只读：写入/更换投递链接前生成“投递链接 X → Y”（items：[{code: 岗位编码, 投递链接}]）。必须原样展示给用户并取得明确确认。"""
    return run_sync('set-apply-link','--preview',items=items)


@operation
def recruiting_apply_link_confirm(items:list[dict],user_quote:str)->dict:
    """用户确认后写入投递链接（只接受岗位所在站点的地址）；不写 Excel。"""
    return run_sync('set-apply-link','--quote',user_quote,items=items)


@operation
def recruiting_site_links_preview(items:list[dict])->dict:
    """只读：写入/更换站点登录链接、账号核对页前生成“当前 → 更新”（items：[{code: 站点编码, 登录链接?, 账号核对页?, retire_to?: 同公司的保留站点}]）。必须原样展示给用户并取得明确确认。"""
    return run_sync('set-site-links','--preview',items=items)


@operation
def recruiting_site_links_confirm(items:list[dict],user_quote:str,expected_sha:str|None=None)->dict:
    """用户确认后写入站点链接（登录链接可为空：该站点只能在岗位页登录）；不写 Excel。"""
    return run_sync('set-site-links','--quote',user_quote,*(['--expect-sha',expected_sha] if expected_sha else []),items=items)


@operation
def recruiting_json_write_excel()->dict:
    """JSON→Excel 唯一写入：标签门禁、状态确认门禁、写出、逐格读回。仅在用户选择“现在写入”后调用。"""
    return run_sync('write')


@operation
def recruiting_sync_pending()->dict:
    """是否存在未写入 Excel 的 JSON 改动，以及用户是否已选择下一轮再写入。"""
    return run_sync('pending')


@operation
def recruiting_write_defer(user_quote:str)->dict:
    """用户选择“下一轮再写入”时登记，绑定当前 JSON 数据摘要。"""
    return run_sync('defer','--quote',user_quote)


@operation
def recruiting_status_check()->dict:
    """校验标签体系、名额收口、JSON 与 Excel 差异。"""
    return run_sync('check')


@operation
def recruiting_data_query(company:str|None=None,sheet:str|None=None,offset:int=0,limit:int=30)->dict:
    """只读：按公司或工作表分页读取主数据 JSON。"""
    args=['query','--offset',str(offset),'--limit',str(limit)]
    if company:args+=['--company',company]
    if sheet:args+=['--sheet',sheet]
    return run_sync(*args)


@operation
def recruiting_company_report(companies:list[str])->dict:
    """只读：按公司汇报主链接、校招计划（含管培/产培）、岗位按投递状态计数与列表、投递名额与占用、剩余名额。"""
    if not companies:raise ValueError('至少指定一家公司')
    args=['report']
    for c in companies:args+=['--company',c]
    return {'reports':run_sync(*args)}


def queue():
    root=os.environ.get('CAMPUS_SEARCH_PRIVATE_DIR')
    if not root:raise ValueError('CAMPUS_SEARCH_PRIVATE_DIR must point to a private workspace directory')
    return Queue(root)


@operation
def search_batch_create(manifest:dict)->dict:
    """Freeze approved company/scope queue. Does not browse or mark facts verified."""
    return queue().create(manifest)


@operation
def search_dispatch(batch_id:str,channels:list[str]|None=None,lease_seconds:int=900,companies_per_channel:int=10)->dict:
    """Assign up to ten companies per browser by default; jobs/plans share their company lease. Host must execute browser calls."""
    from browser_capabilities import select_channels
    root=os.environ.get('CAMPUS_RECRUITING_WORKSPACE')
    if not root:raise ValueError('CAMPUS_RECRUITING_WORKSPACE_REQUIRED')
    selection=select_channels(root,'search',channels)
    return {**queue().dispatch(batch_id,selection['channels'],lease_seconds,companies_per_channel), 'browserSelection':selection}


@operation
def search_result_record(batch_id:str,task_id:str,lease_token:str,evidence_ref:str)->dict:
    """Bind captured evidence to its active company/channel lease; reject stale or duplicate delivery."""
    return queue().record(batch_id,task_id,lease_token,evidence_ref)


@operation
def search_round_check(batch_id:str)->dict:
    """Read current-round counts, missing tasks, blockers and evidence integrity. Not proof of website completeness."""
    return queue().check(batch_id)


@operation
def search_round_advance(batch_id:str)->dict:
    """Advance after every company has evidence and a check outcome; carry unresolved items without marking them passed. Never grants user approval."""
    return queue().advance(batch_id)


@operation
def search_task_retry(batch_id:str,task_id:str,recovery_reason:str)->dict:
    """Explicit recovery of one blocked task; inspect original page before retry to avoid duplicate work."""
    return queue().retry(batch_id,task_id,recovery_reason)



