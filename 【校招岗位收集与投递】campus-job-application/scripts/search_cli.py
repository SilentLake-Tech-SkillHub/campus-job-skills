#!/usr/bin/env python3
"""Independent search-component API. JSON input on stdin, JSON result on stdout."""
import argparse,json,os,sys
from pathlib import Path
from search_service import OPERATIONS,execute
from search_paths import configure

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('operation',choices=['list',*OPERATIONS])
    parser.add_argument('--workspace')
    args=parser.parse_args()
    if args.workspace:
        configure(args.workspace)
    if args.operation=='list':result=[{'name':name,'description':fn.__doc__} for name,fn in OPERATIONS.items()]
    else:result=execute(args.operation,json.load(sys.stdin))
    print(json.dumps(result,ensure_ascii=False))

if __name__=='__main__':
    try:main()
    except Exception as e:print(json.dumps({'error':str(e)},ensure_ascii=False),file=sys.stderr);sys.exit(1)
