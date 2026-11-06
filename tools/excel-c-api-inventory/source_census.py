"""Extract narrowly-scoped declaration candidates from the pinned XLCALL.H.

This is not a general C parser or an independent review of its own output.
"""
import ast
import collections
import hashlib
import json
import re
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
REVISION = '6c8354d6536e18b58559f9cc3134f25a9e58d478'
BLOB = '4eb50e8fd3aea40d283181b7dcc74871ba88d854'
PATH = 'tools/Excel2013XLLSDK/INCLUDE/XLCALL.H'
# Reviewed provenance tuples, extended only with source/version evidence and tests.
# This is a source allowlist, not a duplicate CAPI identity register.
APPROVED_SOURCES = frozenset({('evnekdev/excel-api', REVISION, PATH, BLOB)})

class ExtractionError(ValueError):
    pass

def pinned_header():
    def git(*args):
        p = subprocess.run(['git', *args], cwd=ROOT, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if p.returncode:
            raise ExtractionError(p.stderr.decode('utf8', 'replace').strip())
        return p.stdout
    if git('rev-parse', f'{REVISION}:{PATH}').decode().strip() != BLOB:
        raise ExtractionError('pinned source blob mismatch')
    return git('cat-file', 'blob', BLOB).decode('latin1')

def uncomment(s):
    return re.sub(r'/\*.*?\*/|//[^\n]*', lambda m: re.sub(r'[^\n]', ' ', m.group()), s, flags=re.S)

def value_of(expression, definitions, active=()):
    """Evaluate only integer literals, names, parentheses, unary +/- and bitwise OR."""
    # Python's AST is only a safe evaluation mechanism, not the source grammar.
    # Reject Python-only numeral spellings before parsing (e.g. 1_000, 0o10).
    for token in re.findall(r'\w+|[^\w\s]', expression, flags=re.ASCII):
        if token[0].isdigit() and not re.fullmatch(r'(?:0[xX][0-9a-fA-F]+|0|[1-9][0-9]*)', token):
            raise ExtractionError(f'unsupported integer literal {token!r}')
    try:
        tree = ast.parse(expression, mode='eval')
    except SyntaxError as e:
        raise ExtractionError(f'unsupported expression {expression!r}') from e
    def walk(node):
        if isinstance(node, ast.Expression): return walk(node.body)
        if isinstance(node, ast.Constant) and type(node.value) is int: return node.value
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.UAdd, ast.USub)):
            n = walk(node.operand); return n if isinstance(node.op, ast.UAdd) else -n
        if isinstance(node, ast.BinOp) and isinstance(node.op, ast.BitOr):
            return walk(node.left) | walk(node.right)
        if isinstance(node, ast.Name):
            if node.id in active: raise ExtractionError(f'recursive alias {node.id}')
            if node.id not in definitions: raise ExtractionError(f'unknown identifier {node.id}')
            return value_of(definitions[node.id], definitions, active + (node.id,))
        raise ExtractionError(f'unsupported expression {expression!r}')
    return walk(tree)

def macro_kind(name, expression):
    if name.startswith('xlflow'): return 'flow-code'
    if name.startswith('xlf'): return 'worksheet-function-id'
    if name.startswith('xlc'): return 'macro-command-id'
    if name.startswith('xltype'): return 'type-flag'
    if name.startswith('xlbit'): return 'ownership-flag'
    if name.startswith('xlerr'): return 'error-code'
    if name.startswith('xlret'): return 'return-code'
    if name.startswith('xlevent'): return 'event-id'
    if name in ('xlCommand', 'xlSpecial', 'xlIntl', 'xlPrompt'): return 'function-modifier'
    if re.search(r'\bxlSpecial\b', expression): return 'special-function-id'
    return 'other-constant'

def explicit_alias(expression, definitions):
    """Equal numeric values do not establish an alias."""
    try:
        node = ast.parse(expression, mode='eval').body
    except SyntaxError:
        return None
    return node.id if isinstance(node, ast.Name) and node.id in definitions else None

def macro_subfamily(name):
    for prefix, family in (('xlHpcRet', 'cluster-return-code'), ('xlMode', 'edit-mode'),
                           ('dt', 'document-type'), ('ht', 'mouse-hit-test')):
        if name.startswith(prefix): return family
    if name == 'xlUDF': return 'user-defined-function-id'
    return None

def extract(text):
    clean = uncomment(text)
    lines = clean.splitlines()
    accounted = bytearray(len(clean))
    def cover(start, end):
        accounted[start:end] = b'\x01' * (end - start)
    offset = 0
    guards = {}; stack = []; seen_else = []
    for line, content in enumerate(lines, 1):
        directive = re.match(r'^\s*#\s*(ifdef|ifndef|if|elif|else|endif)\b\s*(.*?)\s*$', content)
        if directive:
            op, arg = directive.groups()
            if op in ('ifdef', 'ifndef'):
                if not re.fullmatch(r'[A-Za-z_]\w*', arg): raise ExtractionError(f'invalid conditional on line {line}')
                stack.append(arg if op == 'ifdef' else '!' + arg)
                seen_else.append(False)
            elif op == 'else':
                if not stack or arg or seen_else[-1]: raise ExtractionError(f'unmatched or repeated else on line {line}')
                seen_else[-1] = True
                stack[-1] = stack[-1][1:] if stack[-1].startswith('!') else '!' + stack[-1]
            elif op == 'endif':
                if not stack or arg: raise ExtractionError(f'unmatched endif on line {line}')
                stack.pop(); seen_else.pop()
            else:
                raise ExtractionError(f'unsupported preprocessor {op} on line {line}')
        if content.lstrip().startswith('#'):
            if not (directive or re.match(r'^\s*#\s*define\b', content) or
                    re.fullmatch(r'\s*#\s*pragma\s+once\s*', content)):
                raise ExtractionError(f'unsupported preprocessor directive on line {line}: {content.strip()}')
            cover(offset, offset + len(content))
        guards[line] = list(stack)
        offset += len(content) + 1
    if stack: raise ExtractionError('unterminated preprocessor conditional')
    candidates = []
    def add(symbol, kind, start, end, **fields):
        if guards[start]: fields['guard'] = guards[start]
        candidates.append(dict(symbol=symbol, kind=kind, line_start=start, line_end=end, **fields))
    macros = []
    for line, content in enumerate(lines, 1):
        m = re.match(r'^\s*#\s*define\s+([A-Za-z_]\w*)(.*)$', content)
        if not m:
            if re.match(r'^\s*#\s*define\b', content):
                raise ExtractionError(f'malformed macro definition on line {line}')
            continue
        if not m.group(2) or not m.group(2)[0].isspace() or content.rstrip().endswith('\\'):
            raise ExtractionError(f'unsupported function-like or multiline macro on line {line}')
        expr = m.group(2).strip()
        if not expr: raise ExtractionError(f'empty macro on line {line}')
        macros.append((line, m.group(1), expr))
    definitions = {}
    for line, name, expr in macros:
        if name in definitions: raise ExtractionError(f'duplicate macro {name} on line {line}')
        definitions[name] = expr
    for line, name, expr in macros:
        details = {'expression':expr, 'value':value_of(expr, definitions, (name,))}
        alias = explicit_alias(expr, definitions)
        if alias: details['alias_of'] = alias
        family = macro_subfamily(name)
        if family: details['subfamily'] = family
        if guards[line]: details['guard'] = guards[line]
        add(name, macro_kind(name, expr), line, line, **details)
    for match in re.finditer(r'(?m)^\s*typedef\s+([A-Za-z_]\w*)\s+([A-Za-z_]\w*)\s*;', clean):
        cover(match.start(), match.end())
        line = clean.count('\n', 0, clean.index('typedef', match.start(), match.end())) + 1
        end_line = clean.count('\n', 0, match.end() - 1) + 1
        add(match.group(2), 'scalar-typedef', line, end_line, underlying=match.group(1))
    for match in re.finditer(r'(?m)^\s*typedef\s+struct\s+([A-Za-z_]\w*)\s*\{', clean):
        depth = 1; pos = match.end()
        while pos < len(clean) and depth:
            depth += (clean[pos] == '{') - (clean[pos] == '}'); pos += 1
        if depth: raise ExtractionError('unterminated struct')
        end = clean.find(';', pos)
        if end < 0: raise ExtractionError('unterminated struct typedef')
        names = clean[pos:end].strip()
        m = re.fullmatch(r'([A-Za-z_]\w*)(?:\s*,\s*\*\s*([A-Za-z_]\w*))?', names)
        if not m: raise ExtractionError(f'unsupported struct typedef names {names!r}')
        cover(match.start(), end + 1)
        start_line = clean.count('\n', 0, clean.index('typedef', match.start(), match.end())) + 1
        end_line = clean.count('\n', 0, end) + 1
        tag, name, pointer = match.group(1), m.group(1), m.group(2)
        fields = {'tag': tag}
        if name == 'XLOPER':
            if not re.search(r'#ifdef\s+__cplusplus\s+WORD\s+xbool\s*;\s*#else\s+WORD\s+bool\s*;\s*#endif', clean[match.start():end]):
                raise ExtractionError('XLOPER conditional member forms changed')
            fields['conditional_fields'] = [{'guard':'__cplusplus', 'declaration':'WORD xbool'},
                                            {'guard':'!__cplusplus', 'declaration':'WORD bool'}]
        add(name, 'struct', start_line, end_line, **fields)
        add(tag, 'struct-tag', start_line, end_line, target=name)
        if pointer: add(pointer, 'pointer-typedef', start_line, end_line, target=name)
    pattern = r'(?m)^\s*typedef\s+(int\s*\(CALLBACK\s*\*\s*([A-Za-z_]\w*)\)\s*\([^;]+\))\s*;'
    for match in re.finditer(pattern, clean):
        cover(match.start(), match.end())
        line = clean.count('\n', 0, clean.index('typedef', match.start(), match.end())) + 1
        end_line = clean.count('\n', 0, match.end() - 1) + 1
        add(match.group(2), 'function-pointer-typedef', line, end_line, signature=' '.join(match.group(1).split()))
    pattern = r'(?m)^\s*((?:int|long)\s+(?:_cdecl|pascal)\s+([A-Za-z_]\w*)\s*\([^;{}]*\))\s*;'
    opening = re.search(r'#ifdef\s+__cplusplus\s+extern\s+"C"\s*\{\s*#endif', clean)
    if not opening:
        raise ExtractionError('conditional C linkage opening changed')
    closing = re.search(r'#ifdef\s+__cplusplus\s*}\s*#endif', clean)
    if not closing or closing.start() <= opening.end():
        raise ExtractionError('conditional C linkage closing changed')
    cover(opening.start(), opening.end())
    cover(closing.start(), closing.end())
    for match in re.finditer(pattern, clean):
        if not opening.end() <= match.start(1) < closing.start():
            raise ExtractionError('function prototype outside conditional C linkage block')
        line = clean.count('\n', 0, match.start(1)) + 1
        cover(match.start(), match.end())
        end_line = clean.count('\n', 0, match.end() - 1) + 1
        add(match.group(2), 'function-prototype', line, end_line, signature=' '.join(match.group(1).split()),
            linkage='extern "C" when __cplusplus')
    for pos, char in enumerate(clean):
        if not accounted[pos] and not char.isspace():
            line = clean.count('\n', 0, pos) + 1
            raise ExtractionError(f'unsupported active source construct on line {line}: {lines[line-1].strip()}')
    keys = [(x['kind'], x['symbol']) for x in candidates]
    if len(keys) != len(set(keys)): raise ExtractionError('duplicate declaration identity')
    return sorted(candidates, key=lambda x: (x['line_start'], x['line_end'], x['kind'], x['symbol']))

def census(text):
    candidates = extract(text)
    counts = dict(sorted(collections.Counter(x['kind'] for x in candidates).items()))
    return {'source': {'repository':'evnekdev/excel-api','revision':REVISION,'path':PATH,'blob_sha':BLOB,
                       'sha256':hashlib.sha256(text.encode('latin1')).hexdigest(),'header_version':'15.0'},
            'complete':True,
            'method':'compare the entire extracted candidate set; completeness attestation is not authenticated',
            'included_categories':list(counts), 'counts':counts,
            'exclusions':['preprocessor control and pragma directives', 'comments', 'nested aggregate members as separate identities'],
            'ambiguities':['authenticated completeness acceptance and complete C semantics are not established'],
            'review':{'status':'generated census is not an authenticated completeness attestation','inspection_date':'2026-11-06'},
            'entries':[dict(path=PATH, disposition='eligible', source_blob=BLOB, **x) for x in candidates]}

if __name__ == '__main__':
    import argparse
    import sys
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', type=Path, help='write the candidate census as UTF-8 JSON')
    args = parser.parse_args()
    try:
        payload = json.dumps(census(pinned_header()), indent=2) + '\n'
        if args.output:
            args.output.write_text(payload, encoding='utf8', newline='\n')
        else:
            print(payload, end='')
    except (ExtractionError, OSError) as e:
        print(f'FAIL: {e}', file=sys.stderr); sys.exit(1)
