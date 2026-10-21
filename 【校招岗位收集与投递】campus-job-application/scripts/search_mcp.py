#!/usr/bin/env python3
"""campus-search MCP：多浏览器搜索任务队列，以及主数据 JSON 的固定读写接口（薄封装 json_excel_sync.py）。"""
import json
import os
import sys
from mcp.server.fastmcp import FastMCP
from search_queue import Queue

server=FastMCP(
    'campus-search',
    instructions=(
        '校招搜索与主数据接口：搜索任务队列（公司租约、证据收回、分轮检查）和主数据 JSON 的固定读写。'
        '不登录、不填表、不提交申请。'
        '主数据是工作区 岗位交接/工作簿全量.json；AI 只读 JSON，不读不改 Excel，Excel 只能由 recruiting_json_write_excel 写出。'
        '用户求职目标与选岗规则只在工作区 产品管理/求职目标.json。'
        '选岗、去留、顺序由用户决定；不得因名额提前剔除任何符合要求的岗位，只展示符合要求的岗位和名额限制。'
        '任何投递状态变更先 recruiting_status_propose，把“当前状态是X，更新状态是Y”原样给用户看，用户确认后才 recruiting_status_confirm。'
    )
)

SYNC_SCRIPT=os.path.join(os.path.dirname(os.path.abspath(__file__)),'json_excel_sync.py')


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


@server.tool()
def recruiting_status_propose(items:list[dict])->dict:
    """只读：生成“当前状态是X，更新状态是Y”（含名额收口派生变更）。必须原样展示给用户并取得明确确认。"""
    return run_sync('propose',items=items)


@server.tool()
def recruiting_status_confirm(items:list[dict],user_quote:str,confirmation_ref:str='')->dict:
    """用户明确确认 propose 展示的变更后调用：记入确认台账并更新主数据 JSON，不写 Excel。"""
    return run_sync('confirm','--quote',user_quote,'--ref',confirmation_ref,items=items)


@server.tool()
def recruiting_rows_upsert(items:list[dict])->dict:
    """按业务键增改主数据 JSON 的非状态字段（新岗位、JD 等）；不能改投递状态，也不能更换已有岗位的链接（用 recruiting_link_change_preview/confirm）。"""
    return run_sync('upsert',items=items)


@server.tool()
def recruiting_link_change_preview(items:list[dict])->dict:
    """只读：更换已有岗位链接前生成“岗位链接 X → Y”（items：[{sheet,company,link,title,new_link}]，link/title 为该行当前值）。必须原样展示给用户并取得明确确认。"""
    return run_sync('relink','--preview',items=items)


@server.tool()
def recruiting_link_change_confirm(items:list[dict],user_quote:str)->dict:
    """用户确认后更换已有岗位链接：投递状态不变、无需重新确认，确认记录与志愿顺序等投递数据随行迁移；不写 Excel。"""
    return run_sync('relink','--quote',user_quote,items=items)


@server.tool()
def recruiting_json_write_excel()->dict:
    """JSON→Excel 唯一写入：标签门禁、状态确认门禁、写出、逐格读回。仅在用户选择“现在写入”后调用。"""
    return run_sync('write')


@server.tool()
def recruiting_sync_pending()->dict:
    """是否存在未写入 Excel 的 JSON 改动，以及用户是否已选择下一轮再写入。"""
    return run_sync('pending')


@server.tool()
def recruiting_write_defer(user_quote:str)->dict:
    """用户选择“下一轮再写入”时登记，绑定当前 JSON 数据摘要。"""
    return run_sync('defer','--quote',user_quote)


@server.tool()
def recruiting_status_check()->dict:
    """校验标签体系、名额收口、JSON 与 Excel 差异。"""
    return run_sync('check')


@server.tool()
def recruiting_data_query(company:str|None=None,sheet:str|None=None,offset:int=0,limit:int=30)->dict:
    """只读：按公司或工作表分页读取主数据 JSON。"""
    args=['query','--offset',str(offset),'--limit',str(limit)]
    if company:args+=['--company',company]
    if sheet:args+=['--sheet',sheet]
    return run_sync(*args)


@server.tool()
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


@server.tool()
def search_batch_create(manifest:dict)->dict:
    """Freeze approved company/scope queue. Does not browse or mark facts verified."""
    return queue().create(manifest)


@server.tool()
def search_dispatch(batch_id:str,channels:list[str]=['chrome','iab'],lease_seconds:int=900,companies_per_channel:int=10)->dict:
    """Assign up to ten companies per browser by default; jobs/plans share their company lease. Host must execute browser calls."""
    return queue().dispatch(batch_id,channels,lease_seconds,companies_per_channel)


@server.tool()
def search_result_record(batch_id:str,task_id:str,lease_token:str,evidence_ref:str)->dict:
    """Bind captured evidence to its active company/channel lease; reject stale or duplicate delivery."""
    return queue().record(batch_id,task_id,lease_token,evidence_ref)


@server.tool()
def search_round_check(batch_id:str)->dict:
    """Read current-round counts, missing tasks, blockers and evidence integrity. Not proof of website completeness."""
    return queue().check(batch_id)


@server.tool()
def search_round_advance(batch_id:str)->dict:
    """Advance after every company has evidence and a check outcome; carry unresolved items without marking them passed. Never grants user approval."""
    return queue().advance(batch_id)


@server.tool()
def search_task_retry(batch_id:str,task_id:str,recovery_reason:str)->dict:
    """Explicit recovery of one blocked task; inspect original page before retry to avoid duplicate work."""
    return queue().retry(batch_id,task_id,recovery_reason)


if __name__=='__main__':
    server.run(transport='stdio')
