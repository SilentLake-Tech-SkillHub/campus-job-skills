#!/usr/bin/env python3
"""MCP transport adapter; all business execution uses the independent core CLI."""
import functools,json,subprocess,sys,os
from pathlib import Path
from mcp.server.fastmcp import FastMCP
from search_service import OPERATIONS
from search_release import source_fingerprint,schema_fingerprint,interface_fingerprint
loaded_interface=interface_fingerprint()
loaded_schema=schema_fingerprint(OPERATIONS)
server=FastMCP(
    'campus-search',
    instructions=(
        '校招搜索与主数据接口：搜索任务队列（公司租约、证据收回、分轮检查）和主数据 JSON 的固定读写。'
        '不登录、不填表、不提交申请。'
        '主数据是工作区 岗位交接/工作簿全量.json；AI 只读 JSON，不读不改 Excel，Excel 只能由 recruiting_json_write_excel 写出。'
        '用户求职目标与选岗规则只在工作区 产品管理/求职目标.json。'
        '选岗、去留、顺序由用户决定；不得因名额提前剔除任何符合要求的岗位，只展示符合要求的岗位和名额限制。'
        '任何投递状态变更先 recruiting_status_propose，把“当前状态是X，更新状态是Y”原样给用户看，用户确认后才 recruiting_status_confirm。'
        '处理投递需求先 recruiting_application_requirements_begin；有参数覆盖写入一次性需求表，无公司参数默认最新十家。'
        'recruiting_application_forward 从表取数，成功组批后立即清空；同回合默认不重复询问，新回合重新初始化。'
    )
)


def adapter(fn):
    @functools.wraps(fn)
    def invoke(**arguments):
        if interface_fingerprint()!=loaded_interface:raise ValueError('SEARCH_INTERFACE_CODE_STALE')
        env={**os.environ,'CAMPUS_SEARCH_EXPECTED_SOURCE':source_fingerprint(),'CAMPUS_SEARCH_EXPECTED_SCHEMA':loaded_schema}
        p=subprocess.run([sys.executable,'-B',str(Path(__file__).with_name('search_cli.py')),fn.__name__],env=env,input=json.dumps(arguments),capture_output=True,text=True,encoding='utf-8')
        if p.returncode:raise ValueError((p.stderr or p.stdout).strip())
        return json.loads(p.stdout)
    return invoke

for fn in OPERATIONS.values():server.tool()(adapter(fn))
if __name__=='__main__':server.run(transport='stdio')
