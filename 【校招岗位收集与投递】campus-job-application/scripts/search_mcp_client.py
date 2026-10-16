#!/usr/bin/env python3
"""Invoke the installed search server over actual MCP stdio, using the official SDK."""
import argparse
import asyncio
import json
import os
from pathlib import Path
import sys
from mcp import ClientSession,StdioServerParameters
from mcp.client.stdio import stdio_client


async def call(request):
    params=StdioServerParameters(command=sys.executable,args=[str(Path(__file__).with_name('search_mcp.py'))],env=dict(os.environ))
    async with stdio_client(params) as (read,write):
        async with ClientSession(read,write) as session:
            await session.initialize()
            if request.get('method')=='tools/list':
                return (await session.list_tools()).model_dump(mode='json')
            result=await session.call_tool(request['tool'],request.get('arguments',{}))
            if result.isError:raise ValueError('Search MCP rejected request: '+str(result.content))
            if result.structuredContent is not None:return result.structuredContent
            if len(result.content)==1 and getattr(result.content[0],'type',None)=='text':
                return json.loads(result.content[0].text)
            return [x.model_dump(mode='json') for x in result.content]


if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--input',required=True)
    args=parser.parse_args()
    request=json.loads(Path(args.input).read_text(encoding='utf-8'))
    print(json.dumps(asyncio.run(call(request)),ensure_ascii=False,indent=2))
