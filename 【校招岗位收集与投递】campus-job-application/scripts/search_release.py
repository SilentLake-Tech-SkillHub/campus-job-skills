"""Installed source identity shared by independent calls and MCP forwarding."""
import hashlib,inspect,json
from pathlib import Path

def source_fingerprint():
    root=Path(__file__).resolve().parent;digest=hashlib.sha256()
    for file in sorted(root.glob('*.py')):
        if file.name.startswith('test_'):continue
        digest.update(file.name.encode());digest.update(b'\0');digest.update(file.read_bytes())
    return digest.hexdigest()

def schema_fingerprint(operations):
    items=[(name,str(inspect.signature(fn))) for name,fn in sorted(operations.items())]
    return hashlib.sha256(json.dumps(items,ensure_ascii=False).encode()).hexdigest()

def interface_fingerprint():
    root=Path(__file__).resolve().parent
    return hashlib.sha256((root/'search_mcp.py').read_bytes()).hexdigest()
