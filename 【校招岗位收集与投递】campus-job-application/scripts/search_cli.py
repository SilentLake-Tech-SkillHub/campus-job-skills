#!/usr/bin/env python3
"""Independent search-component API. JSON input on stdin, JSON result on stdout."""
import argparse,json,os,sys
from pathlib import Path
from search_release import source_fingerprint,schema_fingerprint
loaded_source=source_fingerprint()
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
    else:
        if source_fingerprint()!=loaded_source:raise ValueError('SEARCH_SOURCE_CHANGED_DURING_LOAD')
        expected=os.environ.get('CAMPUS_SEARCH_EXPECTED_SOURCE')
        schema=os.environ.get('CAMPUS_SEARCH_EXPECTED_SCHEMA')
        if expected and expected!=loaded_source:raise ValueError('SEARCH_VERSION_MISMATCH')
        if schema and schema!=schema_fingerprint(OPERATIONS):raise ValueError('SEARCH_INTERFACE_SCHEMA_STALE')
        result=execute(args.operation,json.load(sys.stdin))
    print(json.dumps(result,ensure_ascii=False))

if __name__=='__main__':
    try:main()
    except Exception as e:print(json.dumps({'error':str(e)},ensure_ascii=False),file=sys.stderr);sys.exit(1)
