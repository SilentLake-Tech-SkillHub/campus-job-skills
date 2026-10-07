"""Static complete-package check; does not register skills or access a real ATS."""
import ast
import hashlib
import json
from pathlib import Path
import re
import subprocess
import sys
import tempfile
from PIL import Image

root = Path(sys.argv[1]).resolve()
skills = list(root.rglob('SKILL.md'))
assert skills, 'no skill entrypoints'
for p in skills:
    source = p.read_text()
    assert source.startswith('---\n') and source.count('\n---\n') >= 1, p
    header = source.split('\n---\n', 1)[0]
    assert re.search(r'^name: [a-z0-9-]+$', header, re.M), p
    assert re.search(r'^description: .+', header, re.M), p
links = 0
for p in root.rglob('*.md'):
    for target in re.findall(r'\]\(([^)]+)\)', p.read_text()):
        if target.startswith(('https:', 'http:', '#', '<')):
            continue
        assert (p.parent/target.split('#')[0]).exists(), (p, target)
        links += 1
scripts = list(root.rglob('*.py'))
for p in scripts:
    ast.parse(p.read_text(), filename=str(p))
synchronized = 0
for p in (root/'docs/flow').glob('*.excalidraw'):
    data = json.loads(p.read_text())
    assert data['type'] == 'excalidraw'
    with Image.open(p.with_suffix('.png')) as im:
        sha = im.info.get('excalidraw_sha256')
        if sha:
            assert sha == hashlib.sha256(p.read_bytes()).hexdigest(), p
            synchronized += 1
validators = [p for p in scripts if p.name in {'validate_batch.py', 'validate_research_progress.py', 'validate_assignments.py'}]
with tempfile.TemporaryDirectory() as temp:
    invalid = Path(temp, 'invalid.json'); invalid.write_text('{}')
    for p in validators:
        result = subprocess.run([sys.executable, str(p), str(invalid)], cwd=temp, capture_output=True, text=True)
        assert result.returncode == 1 and not result.stderr, (p, result.stdout, result.stderr)
        json.loads(result.stdout)  # Independent CLI import/loading and controlled invalid-data failure.
print(json.dumps({'passed': True, 'skill_entrypoints': len(skills), 'local_links': links,
                  'python_sources': len(scripts), 'validators_loaded': len(validators),
                  'source_png_pairs_hash_verified': synchronized,
                  'scope': 'complete copied package; static metadata/references and CLI loading, no platform registration or live ATS'}, ensure_ascii=False))
