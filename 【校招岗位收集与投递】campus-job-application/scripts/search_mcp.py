#!/usr/bin/env python3
"""MCP entry point for dual-browser search; private runtime supplied by environment."""
import os
from mcp.server.fastmcp import FastMCP
from search_queue import Queue

server=FastMCP('campus-search',instructions='Dispatch independent company tasks to chrome and iab concurrently through the host browser MCP. Record actual browser evidence. No login, application, filling or submission. Run round checks before advancing; company HTML confirmation remains separate.')


def queue():
    root=os.environ.get('CAMPUS_SEARCH_PRIVATE_DIR')
    if not root:raise ValueError('CAMPUS_SEARCH_PRIVATE_DIR must point to a private workspace directory')
    return Queue(root)


@server.tool()
def search_batch_create(manifest:dict)->dict:
    """Freeze approved company/scope queue. Does not browse or mark facts verified."""
    return queue().create(manifest)


@server.tool()
def search_dispatch(batch_id:str,channels:list[str]=['chrome','iab'],lease_seconds:int=900)->dict:
    """Atomically assign up to two independent tasks; host executes both browser calls concurrently."""
    return queue().dispatch(batch_id,channels,lease_seconds)


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
