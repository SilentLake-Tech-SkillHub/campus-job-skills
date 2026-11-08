"""Resolve declared component contracts without copying their source files."""
import argparse,hashlib,json,os
from pathlib import Path

COMPONENT=Path(__file__).resolve().parents[1]
MANIFEST=COMPONENT/'component-dependencies.json'

def dependency(name):return json.loads(MANIFEST.read_text(encoding='utf-8'))['dependencies'][name]

def contract_path(name):
    spec=dependency(name);explicit=os.environ.get(spec['environment'])
    if explicit:return Path(explicit).expanduser().resolve()
    candidates=[COMPONENT.parent/spec['workspacePath'],COMPONENT.parent/spec['checkoutPath']]
    return next((p for p in candidates if p.is_file()),candidates[0])

def load_contract(name,file=None):
    spec=dependency(name);target=Path(file) if file is not None else contract_path(name)
    try:data=target.read_bytes()
    except OSError as e:raise ValueError('COMPONENT_DEPENDENCY_MISSING: '+name+'; set '+spec['environment']+' to the declared owner file; run component_dependencies.py check') from e
    if hashlib.sha256(data).hexdigest()!=spec['sha256']:raise ValueError('COMPONENT_DEPENDENCY_VERSION_MISMATCH: '+name)
    value=json.loads(data)
    if value.get('schemaVersion')!=spec['schemaVersion']:raise ValueError('COMPONENT_DEPENDENCY_SCHEMA_MISMATCH: '+name)
    return value

def main():
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('action',choices=['describe','check']);args=parser.parse_args();spec=dependency('toolArguments')
    if args.action=='describe':print(json.dumps(spec,ensure_ascii=False))
    else:
        load_contract('toolArguments');print(json.dumps({'status':'verified','path':str(contract_path('toolArguments')),'revision':spec['revision'],'sha256':spec['sha256']},ensure_ascii=False))
if __name__=='__main__':main()
