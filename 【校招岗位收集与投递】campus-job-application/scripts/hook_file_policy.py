"""File-operation inspection shared by the project Hook's independent rules.

Inspects literal tool arguments; it is not a shell/runtime security sandbox.
No command, submitted program, workbook or user document is executed/read here.
"""
import ast
from dataclasses import dataclass
import json
import os
import re
import shlex
from pathlib import Path

from host_payload import call_parts

PATCH_HEADERS = re.compile(r'^\*\*\* (?:(?:Update|Add|Delete) File|Move to): *(.+?)\s*$', re.M)
READ_TOOLS = {'read', 'read_file', 'readfile', 'view_image', 'get_file', 'open_file'}
EDIT_TOOLS = {'write', 'write_file', 'edit', 'edit_file', 'replace', 'apply_patch', 'create_file', 'notebookedit'}
SHELL_TOOLS = {'bash', 'shell', 'exec', 'exec_command', 'execute_command', 'run_command', 'terminal'}
DISPLAY = {'echo', 'printf', 'true', 'false', 'pwd', 'date', 'which', 'command', 'ls', 'stat'}
READ_COMMANDS = {'cat', 'head', 'tail', 'less', 'more', 'xxd', 'hexdump', 'strings', 'file', 'jq', 'wc', 'diff'}
LITERAL = r'''("(?:\\.|[^"\\])*"|'(?:\\.|[^'\\])*')'''
# Which host tools only carry text, and which argument names mean a file: ONE definition, owned by the
# application component next to its own executor hook (sibling folder of this component). Missing or
# unreadable: no tool is text-only and no argument name is trusted; file tools then fall back to the
# registered-file scan in mentioned(), which still protects every registered data file.
TOOL_ARGUMENTS_FILE = Path(__file__).resolve().parents[2] / 'job-application-assistant-v2/hooks/tool-arguments.json'


def tool_arguments(spec_file=TOOL_ARGUMENTS_FILE):
    try:
        spec = json.loads(Path(spec_file).read_text(encoding='utf-8'))
        return frozenset(t.lower() for t in spec['textTools']), tuple(spec['fileArgumentKeys'])
    except (OSError, ValueError, KeyError, TypeError):
        return frozenset(), ()


TEXT_TOOLS, PATH_KEYS = tool_arguments()


def _has_file_argument(value):
    if isinstance(value, dict):
        return any(k in PATH_KEYS or _has_file_argument(v) for k, v in value.items())
    if isinstance(value, list):
        return any(_has_file_argument(v) for v in value)
    return False


def text_only(name, ti):
    return name in TEXT_TOOLS and isinstance(ti, dict) and not _has_file_argument(ti)


@dataclass(frozen=True)
class Effect:
    kind: str                 # read / write / delete
    path: Path
    tree: bool = False
    reason: str = ''


def data_call_parts(payload):
    """Data-policy adapter only. Never changes the component-lock adapter."""
    tool, ti = call_parts(payload)
    text = ti if isinstance(ti, str) else ti.get('command', '') if isinstance(ti, dict) else ''
    if isinstance(text, str) and text.strip().startswith('*** Begin Patch\n') and text.strip().endswith('*** End Patch'):
        return 'apply_patch', text
    if isinstance(ti, str) and ti.lstrip().startswith('{'):
        try:
            decoded = json.loads(ti)
            if isinstance(decoded, dict): ti = decoded
        except (ValueError, TypeError):
            pass
    return tool, ti


def mentioned(payload_value, cwd, registered, reason):
    """Fail closed for unresolved operations involving registered objects only."""
    texts = []
    def visit(value):
        if isinstance(value, str): texts.append(value)
        elif isinstance(value, dict):
            for v in value.values(): visit(v)
        elif isinstance(value, list):
            for v in value: visit(v)
    visit(payload_value)
    result = []
    for raw in registered:
        target = Path(raw.rstrip('/')).resolve(strict=False)
        for text in texts:
            if '://' in text and not any(ch.isspace() for ch in text): continue
            candidate = path_at(text, cwd)
            if candidate == target or raw.endswith('/') and candidate is not None and target in candidate.parents:
                result.append(Effect('unknown', candidate, reason=reason)); continue
            # Absolute and cwd-relative references, including quoted paths
            # with spaces. Matching a basename in another directory is not a hit.
            import os
            refs = [str(target), os.path.relpath(target, cwd)]
            tail = r'(?:/[^\s\'";|&]*)?' if raw.endswith('/') else ''
            hit = any(re.search(r'(?<![\w./-])' + re.escape(ref) + tail + r'(?![\w./-])', text) for ref in refs)
            if not hit:
                literals = re.findall(LITERAL, text)
                for value in literals:
                    candidate = path_at(_literal(value), cwd)
                    if candidate == target or raw.endswith('/') and candidate is not None and target in candidate.parents:
                        result.append(Effect('unknown', candidate, reason=reason)); break
            else:
                result.append(Effect('unknown', target, reason=reason))
    return list(dict.fromkeys(result))


def path_at(value, cwd):
    if not isinstance(value, str) or not value or '://' in value or value.startswith('-'):
        return None
    p = Path(value).expanduser()
    return (p if p.is_absolute() else Path(cwd) / p).resolve(strict=False)


def touches(effect, target):
    target = Path(target).resolve(strict=False)
    return effect.path == target or effect.tree and effect.path in target.parents


def _literal(value):
    try:
        result = ast.literal_eval(value)
        return result if isinstance(result, str) else None
    except (ValueError, SyntaxError):
        return None


def _outside_strings(program):
    outside, quote, escaped = [], None, False
    for ch in program:
        outside.append(quote is None)
        if escaped: escaped = False; continue
        if quote and ch == '\\': escaped = True; continue
        if quote is None and ch in "'\"`": quote = ch
        elif quote == ch: quote = None
    return outside


def _command_calls(code):
    """Read literal command objects without evaluating JavaScript."""
    for match in re.finditer(r'tools\.exec_command\s*\(\s*\{', code):
        start, i, depth, quote, escaped = match.start(), match.end(), 1, None, False
        while i < len(code) and depth:
            ch = code[i]
            if escaped: escaped = False
            elif quote and ch == '\\': escaped = True
            elif quote == ch: quote = None
            elif quote is None and ch in "'\"`": quote = ch
            elif quote is None and ch == '{': depth += 1
            elif quote is None and ch == '}': depth -= 1
            i += 1
        if depth: continue
        body = code[match.end():i - 1]
        outside = _outside_strings(body)
        def property_match(keys):
            pattern = r'(?:^|,)\s*[\'\"]?(?:' + keys + r')[\'\"]?\s*:\s*' + LITERAL
            return next((m for m in re.finditer(pattern, body, re.S) if outside[m.start()]), None)
        cmd = property_match('cmd|command')
        directory = property_match('cwd|workdir')
        if cmd:
            yield start, i, _literal(cmd.group(1)), _literal(directory.group(1)) if directory else None


def _inline(program, language, cwd):
    out = []
    def add(kind, value):
        p = path_at(value, cwd)
        if p is not None:
            out.append(Effect(kind, p))
    if language.startswith('python'):
        try:
            tree = ast.parse(program)
        except SyntaxError:
            tree = None
        for node in ast.walk(tree) if tree else ():
            if not isinstance(node, ast.Call):
                continue
            func = node.func
            name = func.id if isinstance(func, ast.Name) else func.attr if isinstance(func, ast.Attribute) else ''
            args = [a.value if isinstance(a, ast.Constant) else None for a in node.args]
            if name == 'open' and (args or any(k.arg == 'file' for k in node.keywords)):
                filename = args[0] if args else next((k.value.value for k in node.keywords if k.arg == 'file' and isinstance(k.value, ast.Constant)), None)
                mode = args[1] if len(args) > 1 else next((k.value.value for k in node.keywords if k.arg == 'mode' and isinstance(k.value, ast.Constant)), 'r')
                add('write' if isinstance(mode, str) and any(c in mode for c in 'wax+') else 'read', filename)
            elif name in {'read_excel', 'load_workbook', 'read_csv'} and args:
                add('read', args[0])
            elif name in {'to_excel', 'save', 'write_file'} and args:
                add('write', args[0])
            elif name in {'replace', 'rename', 'copyfile', 'copy', 'copy2'} and len(args) > 1:
                add('read', args[0]); add('write', args[1])
            elif name in {'remove', 'unlink'} and args:
                add('delete', args[0])
            elif isinstance(func, ast.Attribute) and isinstance(func.value, ast.Call):
                owner = func.value
                if isinstance(owner.func, ast.Name) and owner.func.id == 'Path' and owner.args and isinstance(owner.args[0], ast.Constant):
                    if name == 'open':
                        mode = args[0] if args else next((k.value.value for k in node.keywords if k.arg == 'mode' and isinstance(k.value, ast.Constant)), 'r')
                        add('write' if isinstance(mode, str) and any(c in mode for c in 'wax+') else 'read', owner.args[0].value)
                    elif name in {'read_text', 'read_bytes'}:
                        add('read', owner.args[0].value)
                    elif name in {'write_text', 'write_bytes', 'touch', 'unlink'}:
                        add('delete' if name == 'unlink' else 'write', owner.args[0].value)
                    elif name in {'rename', 'replace'} and args:
                        add('delete', owner.args[0].value); add('write', args[0])
        return out
    # Literal fs APIs in inline JS and opaque orchestration wrappers. Never eval.
    pattern = r'\b(readFile(?:Sync)?|writeFile(?:Sync)?|appendFile(?:Sync)?|createWriteStream|unlink(?:Sync)?|write_text|write_bytes)\s*\(\s*' + LITERAL
    outside = _outside_strings(program)
    for match in re.finditer(pattern, program):
        if not outside[match.start()]: continue
        name, raw = match.groups()
        add('read' if name.startswith('read') else 'delete' if name.startswith('unlink') else 'write', _literal(raw))
    for m in re.finditer(r'\b(rename(?:Sync)?|copyFile(?:Sync)?|cpSync)\s*\(\s*' + LITERAL + r'\s*,\s*' + LITERAL, program):
        if not outside[m.start()]: continue
        name, src, dst = m.groups()
        add('delete' if name.startswith('rename') else 'read', _literal(src)); add('write', _literal(dst))
    for m in re.finditer(r'\bopenSync\s*\(\s*' + LITERAL + r'\s*,\s*' + LITERAL, program):
        if not outside[m.start()]: continue
        path, mode = map(_literal, m.groups())
        add('write' if mode and any(c in mode for c in 'wax+') else 'read', path)
    return out


SHELL_VAR = re.compile(r'\$(?:\{([A-Za-z_]\w*)\}|([A-Za-z_]\w*))')


def expand_vars(text, assigned):
    """Expand variables this command assigned or the environment defines.

    One unknown variable leaves the text as written, so callers keep judging it
    the way they judged an unresolved path before.
    """
    missing = False
    def value(match):
        nonlocal missing
        name = match.group(1) or match.group(2)
        found = assigned[name] if name in assigned else os.environ.get(name)
        if not found:
            missing = True
            return match.group(0)
        return found
    expanded = SHELL_VAR.sub(value, text)
    return text if missing else expanded


def note_assignments(words, assigned):
    """A segment of bare assignments lasts for the rest of the command; `X=1 cmd` does not."""
    words = words[1:] if words[:1] == ['export'] else words
    if words and all(re.match(r'^[A-Za-z_]\w*=', w) for w in words):
        for word in words:
            name, raw = word.split('=', 1)
            value = expand_vars(raw, assigned)
            # A computed value ($(...), an unknown variable) stays unknown and hides any same-named env value.
            assigned[name] = '' if '$' in value or '`' in value else value


HEREDOC = re.compile(r"""<<(-?)[ \t]*(?:'([^'\n]+)'|"([^"\n]+)"|\\?([^\s;&|<>()'"]+))""")
SHELL_SINK = re.compile(r'(?:^|[;&|(]\s*|\bsudo\s+|\benv\s+)(?:ba|z)?sh\b')
RUNS_TEXT = re.compile(r'\b(?:sh|bash|zsh)\s+-\w*c\b|\beval\b|\bawk\b|system\(|popen|subprocess|\bexec(?:Sync)?\(|\bspawn')
REDIRECT_OP = re.compile(r'>{1,2}')
DATA_TARGET = re.compile(r'[|&]?[ \t]*([^\s;&|]*)')  # `>|` and `>&` still name a target
WRITE_REDIRECTS = {'>', '>>', '>|', '&>', '&>>', '>&'}
# shlex drops quotes, so `grep '>|' f` would otherwise lex like a real redirect.
QUOTED_OPERATOR = re.compile(r"""(?<![^\s;&|(])(['"])(?:&>>?|>>?|>\||>&)\1(?![^\s;&|)])""")


def _closing_paren(text, opening):
    depth, i = 0, opening
    while i < len(text):
        if text[i] == '(': depth += 1
        elif text[i] == ')':
            depth -= 1
            if depth == 0: return i
        i += 1
    return None


def shell_data_spans(command, base=0):
    """Where the shell passes text along instead of running it: (start, end, kind).

    Kinds: quote, heredoc, comment, compare. Returns None when the command does
    not parse, so callers fall back to treating all of it as code.
    """
    spans, pending, i, n = [], [], 0, len(command)
    while i < n:
        ch = command[i]
        at_word = i == 0 or command[i - 1] in ' \t\n;&|('
        if ch == '\\':
            i += 2
        elif ch == "'":
            end = command.find("'", i + 1)
            if end < 0: return None
            spans.append((base + i, base + end, 'quote')); i = end + 1
        elif ch == '"':
            start, j = i, i + 1
            while j < n and command[j] != '"':
                if command[j] == '\\':
                    j += 2
                elif command.startswith('$(', j) and not command.startswith('$((', j):
                    end = _closing_paren(command, j + 1)
                    inner = None if end is None else shell_data_spans(command[j + 2:end], base + j + 2)
                    if inner is None: return None
                    spans.append((base + start, base + j, 'quote')); spans += inner
                    start = j = end + 1
                else:
                    j += 1
            if j >= n: return None
            spans.append((base + start, base + j, 'quote')); i = j + 1
        elif ch == '#' and at_word:
            end = command.find('\n', i)
            end = n if end < 0 else end
            spans.append((base + i, base + end, 'comment')); i = end
        elif command.startswith('((', i) or command.startswith('[[', i) and at_word:
            end = command.find('))' if ch == '(' else ']]', i + 2)
            if end >= 0:
                spans.append((base + i, base + end, 'compare')); i = end + 2
            else:
                i += 2
        elif command.startswith('<<', i) and not command.startswith('<<<', i) and (m := HEREDOC.match(command, i)):
            line_start = command.rfind('\n', 0, i) + 1
            line_end = command.find('\n', i)
            line = command[line_start:n if line_end < 0 else line_end]
            pending.append((m.group(2) or m.group(3) or m.group(4), bool(m.group(1)), bool(SHELL_SINK.search(line))))
            i = m.end()
        elif ch == '\n' and pending:
            i += 1
            for delimiter, dashed, runs_as_shell in pending:
                start = i
                while i < n:
                    end = command.find('\n', i)
                    end = n if end < 0 else end
                    line = command[i:end]
                    if (line.lstrip('\t') if dashed else line) == delimiter:
                        break
                    i = end + 1
                if not runs_as_shell:
                    spans.append((base + start, base + min(i, n), 'heredoc'))
                end = command.find('\n', i)
                i = n if end < 0 else end + 1
            pending = []
        else:
            i += 1
    return spans


def blank_data_redirects(command, anchors=None, kinds=None):
    """Blank each `>` in data that cannot be a redirect, so `i>=0` and `a > b` stop reading as writes.

    Text in data can still reach a shell (`os.system('echo > x')`), so a `>` stays
    when its target is a named path. With `anchors`, a relative target must also
    name one of them or the command must contain something that runs text.
    """
    spans = shell_data_spans(command)
    if not spans:
        return command
    chars, runs_text = list(command), bool(RUNS_TEXT.search(command))
    def stays(op, end):
        before = command[op.start() - 1] if op.start() else ''
        target = DATA_TARGET.match(command, op.end(), end).group(1)
        bare = target.strip('\'"\\)]},(:')
        if command[op.end():op.end() + 1] == '=' or before in ('=', '-', '<') or not bare:
            return False
        pathlike = '/' in bare or bool(re.search(r'\.[A-Za-z]\w*$', bare))
        if anchors is None:
            return pathlike
        if any(a.rstrip('/') in bare for a in anchors) or bare[0] in '/~' or bare[0] == '$' and '/' in bare:
            return True
        return runs_text and (pathlike or target.lstrip('\\')[:1] in ("'", '"'))
    for start, end, kind in spans:
        if kinds is None or kind in kinds:
            for op in REDIRECT_OP.finditer(command, start, end):
                if not stays(op, end):
                    chars[op.start():op.end()] = ' ' * (op.end() - op.start())
    return ''.join(chars)


def _subcommands(command):
    """Observe executable substitutions, ignoring text inside single quotes."""
    out, i, quote = [], 0, None
    while i < len(command):
        ch = command[i]
        if ch == '\\' and quote != "'":
            i += 2; continue
        if ch in "'\"":
            if quote is None: quote = ch
            elif quote == ch: quote = None
            i += 1; continue
        if quote != "'" and command.startswith('$(', i):
            start, depth, j = i + 2, 1, i + 2
            while j < len(command) and depth:
                if command[j] == '(': depth += 1
                elif command[j] == ')': depth -= 1
                j += 1
            if depth == 0:
                out.append(command[start:j - 1]); i = j; continue
        if ch == '`' and quote != "'":
            end = command.find('`', i + 1)
            if end >= 0:
                out.append(command[i + 1:end]); i = end + 1; continue
        i += 1
    return out




def _find_effects(args, cwd, registered):
    """Listing names is metadata; deleting/executing is not a read-only find."""
    roots, i = [], 0
    while i < len(args):
        arg = args[i]
        if arg in {'-H', '-L', '-P', '-X', '-s', '-d'} or re.fullmatch(r'-O\d+', arg):
            i += 1; continue
        if arg == '-D':
            i += 2; continue
        if arg == '-f' and i + 1 < len(args):
            roots.append(args[i + 1]); i += 2; continue
        if arg.startswith('-') or arg in {'!', '('}: break
        roots.append(arg); i += 1
    paths = [p for value in (roots or ['.']) if (p := path_at(value, cwd)) is not None]
    out = []
    if '-delete' in args:
        out.extend(Effect('delete', p, tree=True) for p in paths)
    actions = {'-exec', '-execdir', '-ok', '-okdir'}.intersection(args)
    if actions:
        reason = 'find 含执行动作 ' + '、'.join(sorted(actions)) + '，无法确定每个匹配项的实际文件操作'
        out.extend(Effect('unknown', p, tree=True, reason=reason) for p in paths)
        # An invoked command can touch a literal file outside find's roots.
        out.extend(mentioned(args, cwd, registered, reason))
    for i, arg in enumerate(args):
        if arg in {'-fprint', '-fprint0', '-fprintf', '-fls', '-files0-from'} and i + 1 < len(args):
            p = path_at(args[i + 1], cwd)
            if p is not None:
                out.append(Effect('read' if arg == '-files0-from' else 'write', p))
    return out


def _read_operands(name, args):
    paths = [a for a in args if not a.startswith('-')]
    file_options = {'jq': {'--from-file'}, 'wc': {'--files0-from'},
                    'diff': {'--from-file', '--to-file'}}.get(name, set())
    paths.extend(a.split('=', 1)[1] for a in args if '=' in a and a.split('=', 1)[0] in file_options)
    return paths


def shell_effects(command, cwd, root, registered=(), sync_script=None):
    out = []
    command = QUOTED_OPERATOR.sub('_', blank_data_redirects(command, kinds={'heredoc'}))
    for sub in _subcommands(command): out.extend(shell_effects(sub, cwd, root, registered, sync_script))
    try:
        lexer = shlex.shlex(command, posix=True, punctuation_chars=';&|<>\n')
        lexer.whitespace = ' \t\r'
        lexer.whitespace_split = True
        tokens = list(lexer)
    except ValueError:
        return mentioned(command, cwd, registered, 'Shell 命令解析失败')
    segments, part = [], []
    for tok in tokens:
        if tok in {';', '&&', '||', '|', '\n', '&'}:
            if part: segments.append(part)
            part = []
        else: part.append(tok)
    if part: segments.append(part)
    assigned = {}
    for segment in segments:
        segment = [expand_vars(tok, assigned) for tok in segment]
        note_assignments(segment, assigned)
        clean, i = [], 0
        while i < len(segment):
            tok = segment[i]
            if (tok in WRITE_REDIRECTS or tok in {'<', '<<'}) and i + 1 < len(segment):
                value = segment[i + 1]
                # The lexer splits `2>x` into `2`, `>`, `x`; left in place, `2` would
                # read as the last argument and hide the real target of `cp a b 2>x`.
                if clean and clean[-1].isdigit(): clean.pop()
                closes_descriptor = tok == '>&' and value == '-'
                if tok != '<<' and value not in {'&', '/dev/null'} and not value.isdigit() and not closes_descriptor:
                    p = path_at(value, cwd)
                    if p is not None: out.append(Effect('read' if tok == '<' else 'write', p))
                i += 2
            else:
                clean.append(tok); i += 1
        while clean and (re.match(r'^\w+=', clean[0]) or clean[0] in {'env', 'command', 'sudo'}): clean.pop(0)
        if not clean: continue
        name, args = Path(clean[0]).name, clean[1:]
        plain = [a for a in args if not a.startswith('-')]
        if name in {'cd', 'pushd'} and plain:
            cwd = path_at(plain[0], cwd); continue
        if name == 'git':
            # Data rules deliberately do not govern ANY Git subcommand.
            # The unchanged component lock makes its own Git decision.
            continue
        if name in {'npm', 'pnpm', 'yarn', 'gh'}:
            # Package mutation remains governed by the component lock. Free
            # script names and issue text are not filesystem path arguments.
            continue
        if name in {'sh', 'bash', 'zsh'} and '-c' in args and args.index('-c') + 1 < len(args):
            out.extend(shell_effects(args[args.index('-c') + 1], cwd, root, registered, sync_script)); continue
        if name == 'perl' and any(a.startswith(('-i', '-pi')) for a in args):
            out += [Effect('write', p) for a in plain[1:] if (p := path_at(a, cwd)) is not None]
            continue
        if name.startswith(('python', 'node', 'ruby', 'perl')):
            if name.startswith('node') and '--test' in args:
                continue
            if name.startswith('python') and '-m' in args and args.index('-m') + 1 < len(args) and args[args.index('-m') + 1] in {'unittest', 'pytest', 'py_compile', 'compileall'}:
                continue
            inline_flag = '-c' if name.startswith('python') else '-e'
            if inline_flag in args and args.index(inline_flag) + 1 < len(args):
                program = args[args.index(inline_flag) + 1]
                inspected = _inline(program, name, cwd)
                out.extend(inspected)
                # Unknown/dynamic programs touching a registered object are
                # refused rather than guessed safe. A plain print is textual.
                textual = re.fullmatch(r'\s*(?:print|console\.log)\s*\(.*\)\s*;?\s*', program, re.S)
                if not inspected and not textual:
                    out.extend(mentioned(program, cwd, registered, '内联程序未能确定文件操作'))
                continue
            script = next((a for a in args if not a.startswith('-')), None)
            if re.fullmatch(r'python(?:\d+(?:\.\d+)*)?', name) and script and sync_script and path_at(script, cwd) == Path(sync_script).resolve():
                continue
            # The component owns this fixed profile writer. An identically named
            # external script is not trusted. Redirections were inspected above.
            profile_writer = Path(root) / '.agents/skills/job-application-assistant-v2/scripts/business-authorization.mjs'
            if name == 'node' and script and path_at(script, cwd) == profile_writer.resolve():
                continue
            for arg in plain[1:]:
                p = path_at(arg, cwd)
                if p is not None: out.append(Effect('unknown', p))
            continue
        if name == 'find':
            out.extend(_find_effects(args, cwd, registered)); continue
        if name in DISPLAY: continue
        if name in READ_COMMANDS:
            out += [Effect('read', p) for a in _read_operands(name, args) if (p := path_at(a, cwd)) is not None]
        elif name in {'cp', 'mv', 'rsync', 'ditto', 'install'} and len(plain) >= 2:
            recursive = any(a in {'-r', '-R', '-a', '--recursive'} for a in args)
            destination = path_at(plain[-1], cwd)
            for a in plain[:-1]:
                p = path_at(a, cwd)
                if p is not None:
                    tree = p.is_dir() or recursive
                    out.append(Effect('delete' if name == 'mv' else 'read', p, tree))
                    if destination is not None:
                        target = destination / p.name if destination.is_dir() else destination
                        out.append(Effect('write', target, tree))
        elif name in {'rm', 'rmdir', 'unlink', 'touch', 'mkdir', 'truncate', 'chmod', 'chown', 'tee'}:
            kind = 'delete' if name in {'rm', 'rmdir', 'unlink'} else 'write'
            recursive = any(a.startswith('-') and not a.startswith('--') and any(c in a[1:] for c in 'rR') or a == '--recursive' for a in args)
            out += [Effect(kind, p, recursive) for a in plain if (p := path_at(a, cwd)) is not None]
        elif name in {'sed', 'perl'}:
            kind = 'write' if any(a.startswith(('-i', '-pi')) for a in args) else 'read'
            out += [Effect(kind, p) for a in plain[1:] if (p := path_at(a, cwd)) is not None]
        elif name in {'rg', 'grep'}:
            names = plain if any(a in args for a in ['-e', '--regexp']) else plain[1:]
            out += [Effect('read', p) for a in names if (p := path_at(a, cwd)) is not None]
        else:
            out.extend(mentioned(clean, cwd, registered, '无法识别涉及登记文件的命令'))
    return out


def file_effects(payload, root, registered=(), sync_script=None):
    tool, ti = data_call_parts(payload)
    # Preserve compound tool names such as exec_command/read_file.
    name = tool.rsplit('__', 1)[-1].rsplit('.', 1)[-1].lower()
    cwd = Path(payload.get('cwd') or root).resolve()
    if name in SHELL_TOOLS and isinstance(ti, dict):
        directory = ti.get('workdir') or ti.get('cwd')
        if isinstance(directory, str): cwd = path_at(directory, cwd)
    out = []
    if text_only(name, ti):
        return []
    paths = [str(ti[k]) for k in PATH_KEYS if isinstance(ti, dict) and isinstance(ti.get(k), str)]
    if name == 'apply_patch':
        text = ti if isinstance(ti, str) else next((v for v in ti.values() if isinstance(v, str) and v.strip().startswith('*** Begin Patch')), '')
        paths = PATCH_HEADERS.findall(text)
    if name in EDIT_TOOLS or any(t in name for t in ['write_file', 'edit_file', 'apply_patch']):
        return [Effect('write', p) for value in paths if (p := path_at(value, cwd)) is not None] or mentioned(ti, cwd, registered, '写入目标参数不完整或解析失败')
    if name in READ_TOOLS:
        return [Effect('read', p) for value in paths if (p := path_at(value, cwd)) is not None] or mentioned(ti, cwd, registered, '读取目标参数不完整或解析失败')
    if paths and name not in SHELL_TOOLS:
        return [Effect('unknown', p, reason='无法识别该文件工具的操作') for value in paths if (p := path_at(value, cwd)) is not None]
    if name in SHELL_TOOLS or name in {'exec_command', 'execute_command'}:
        command = next((ti[k] for k in ['command', 'cmd'] if isinstance(ti, dict) and isinstance(ti.get(k), str)), None)
        if command is not None: return shell_effects(command, cwd, root, registered, sync_script)
    if name == 'exec':
        code = ti if isinstance(ti, str) else next((ti[k] for k in ['input', 'code', 'source'] if isinstance(ti, dict) and isinstance(ti.get(k), str)), '')
        calls = list(_command_calls(code)); remaining = list(code)
        for start, end, command, directory in calls:
            if command is not None:
                out.extend(shell_effects(command, path_at(directory, cwd) if directory else cwd, root, registered, sync_script))
                remaining[start:end] = ' ' * (end - start)
        rest = ''.join(remaining)
        out.extend(_inline(rest, 'js', cwd))
        if calls and len(calls) == len(re.findall(r'tools\.exec_command\s*\(', code)):
            return out
        if re.fullmatch(r'\s*(?:text|notify)\s*\(\s*' + LITERAL + r'\s*\)\s*;?\s*', code, re.S): return out
        if out or not re.search(r'tools\.|\b(?:open|write|read|exec)\b', code): return out
    return out or mentioned(ti, cwd, registered, '无法识别工具或解析出实际文件操作')
