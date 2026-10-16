#!/usr/bin/env python3
"""MCP entry point for dual-browser search; private runtime supplied by environment."""
import os
from mcp.server.fastmcp import FastMCP
from search_queue import Queue
from recruiting_data import RecruitingData

server=FastMCP('campus-search',instructions='Coordinate the recruiting research and data workflow: dual-browser search, private structured JSON drafts, version-bound Excel updates/readback, complete JSON and fixed HTML review exports, explicit user confirmation and handoff checks. No login, application, filling or submission. Initialization, research completion and user confirmation are distinct. Respect the approved batch write timing.')


def data_store():
    workspace=os.environ.get('CAMPUS_RECRUITING_WORKSPACE')
    private=os.environ.get('CAMPUS_SEARCH_PRIVATE_DIR')
    if not workspace or not private:
        raise ValueError('数据接口需要配置CAMPUS_RECRUITING_WORKSPACE及私有目录')
    return RecruitingData(workspace,private)


@server.tool()
def recruiting_data_open(workbook:str,batch:str)->dict:
    """Quickly expand workbook/JSON status and structured schema; no Excel mutation."""
    return data_store().open(workbook,batch)


@server.tool()
def recruiting_data_query(workbook:str,company:str|None=None,sheet:str|None=None,offset:int=0,limit:int=30)->dict:
    """Read a company or sheet from synchronized full JSON with pagination."""
    return data_store().query(workbook,company,sheet,offset,limit)


@server.tool()
def recruiting_draft_put(workbook:str,batch:str,data:dict)->dict:
    """Save an immutable structured exchange draft; does not mark facts or user approval verified."""
    return data_store().draft_put(workbook,batch,data)


@server.tool()
def recruiting_draft_get(workbook:str,batch:str,draft_ref:str,draft_sha256:str)->dict:
    """Read the exact hashed draft bound to this workbook and batch."""
    return data_store().draft(workbook,batch,draft_ref,draft_sha256)


@server.tool()
def recruiting_data_initialize(workbook:str,batch:str,sheet:str,scope:str,expected_sha256:str,company:str|None=None)->dict:
    """Initialize research tables through fixed code only at the batch-authorized write stage."""
    return data_store().initialize(workbook,batch,sheet,scope,expected_sha256,company)


@server.tool()
def recruiting_data_commit(workbook:str,batch:str,draft_ref:str,draft_sha256:str,expected_sha256:str)->dict:
    """Validate and commit a draft to Excel, read back and refresh full/batch JSON and fixed HTML."""
    return data_store().commit(workbook,batch,draft_ref,draft_sha256,expected_sha256)


@server.tool()
def recruiting_review_refresh(workbook:str,batch:str,expected_sha256:str)->dict:
    """Export all existing batches and fixed HTML from current Excel; never import an old snapshot."""
    return data_store().refresh(workbook,batch,expected_sha256)


@server.tool()
def recruiting_review_confirm(workbook:str,batch:str,company_key:str,evidence_ref:str,html_sha256:str,expected_sha256:str)->dict:
    """Record existing explicit user approval of exact HTML; enforce research/ranking evidence. No application authority."""
    return data_store().confirm(workbook,batch,company_key,evidence_ref,html_sha256,expected_sha256)


@server.tool()
def recruiting_handoff_check(workbook:str,batch:str,url:str|None=None)->dict:
    """Check Excel/full JSON/batch JSON/HTML and user approval before handoff; never submit."""
    return data_store().check(workbook,batch,url)


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
