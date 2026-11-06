"""Complementary lexical accounting of the pinned header, independent of the catalogue.

Uses a separate comment state machine and source-token scan. A matching result
is a mechanical cross-check, not independent human or ABI review.
"""
import argparse
import json
import re
import subprocess
import sys
from collections import Counter

REV = '6c8354d6536e18b58559f9cc3134f25a9e58d478'
PATH = 'tools/Excel2013XLLSDK/INCLUDE/XLCALL.H'

def strip_comments(text):
    out = []; i = 0; mode = 'code'
    while i < len(text):
        ch = text[i]; two = text[i:i+2]
        if mode == 'code' and two in ('/*', '//'):
            mode = 'block' if two == '/*' else 'line'; out.extend('  '); i += 2; continue
        if mode == 'block' and two == '*/':
            mode = 'code'; out.extend('  '); i += 2; continue
        if mode == 'line' and ch == '\n': mode = 'code'
        out.append(ch if mode == 'code' or ch == '\n' else ' ')
        i += 1
    if mode == 'block': raise ValueError('unterminated block comment')
    return ''.join(out)

def scan(text):
    clean = strip_comments(text); lines = clean.splitlines(); directives = Counter()
    macros = {}; declarations = set(); typedef_starts = 0; prototypes = 0; accounted = set()
    for i, line in enumerate(lines, 1):
        m = re.match(r'^\s*#\s*(\w+)\b(.*)', line)
        if m:
            accounted.add(i)
            directives[m.group(1)] += 1
            if m.group(1) == 'define':
                n = re.match(r'\s+([A-Za-z_]\w*)', m.group(2))
                if not n: raise ValueError(f'unrecognized define line {i}')
                macros[(n.group(1), i)] = m.group(2)[n.end():].strip()
        m = re.match(r'^\s*(?:int|long)\s+(?:_cdecl|pascal)\s+(\w+)\s*\(', line)
        if m:
            prototypes += 1; accounted.add(i); declarations.add(('function-prototype', m.group(1), i, i))
        if not re.match(r'^\s*typedef\b', line): continue
        typedef_starts += 1
        chunk = []; depth = 0; end = None
        for j in range(i-1, len(lines)):
            current = lines[j]; chunk.append(current)
            for c in current:
                if c == '{': depth += 1
                elif c == '}': depth -= 1
                elif c == ';' and depth == 0: end = j+1; break
            if end: break
        if end is None: raise ValueError(f'unterminated typedef at {i}')
        accounted.update(range(i, end + 1))
        statement = '\n'.join(chunk)
        aggregate = re.match(r'\s*typedef\s+struct\s+(\w+)\s*\{', statement)
        if aggregate:
            tail = re.search(r'}\s*(\w+)(?:\s*,\s*\*\s*(\w+))?\s*;\s*$', statement)
            if not tail: raise ValueError(f'unrecognized aggregate typedef at {i}')
            declarations.add(('struct-tag', aggregate.group(1), i, end))
            declarations.add(('struct', tail.group(1), i, end))
            if tail.group(2): declarations.add(('pointer-typedef', tail.group(2), i, end))
        else:
            callback = re.search(r'\(\s*CALLBACK\s*\*\s*(\w+)\s*\)', statement)
            scalar = re.fullmatch(r'\s*typedef\s+\w+\s+(\w+)\s*;\s*', statement)
            if callback: declarations.add(('function-pointer-typedef', callback.group(1), i, end))
            elif scalar: declarations.add(('scalar-typedef', scalar.group(1), i, end))
            else: raise ValueError(f'unrecognized typedef at {i}')
    unaccounted = [(i, line.strip()) for i, line in enumerate(lines, 1)
                   if i not in accounted and line.strip() and line.strip() not in ('extern "C" {', '}')]
    return macros, declarations, directives, typedef_starts, prototypes, unaccounted

def independent_value(expression, definitions, active=()):
    """Small recursive-descent integer/OR parser, independent of the AST evaluator."""
    tokens = re.findall(r'0[xX][0-9A-Fa-f]+|\d+|[A-Za-z_]\w*|[()|+\-]', expression)
    if ''.join(tokens) != re.sub(r'\s+', '', expression):
        raise ValueError(f'unsupported expression: {expression}')
    index = 0
    def primary():
        nonlocal index
        if index >= len(tokens): raise ValueError('incomplete expression')
        token = tokens[index]; index += 1
        if token in ('+', '-'):
            value = primary(); return value if token == '+' else -value
        if token == '(':
            value = bit_or()
            if index >= len(tokens) or tokens[index] != ')': raise ValueError('unclosed parenthesis')
            index += 1; return value
        if re.fullmatch(r'0[xX][0-9A-Fa-f]+', token): return int(token, 16)
        if re.fullmatch(r'\d+', token):
            if len(token) > 1 and token.startswith('0'):
                if not re.fullmatch(r'0[0-7]+', token): raise ValueError('invalid octal literal')
                return int(token, 8)
            return int(token, 10)
        if token in active: raise ValueError(f'cyclic definition: {token}')
        if token not in definitions: raise ValueError(f'unknown definition: {token}')
        return independent_value(definitions[token], definitions, active + (token,))
    def bit_or():
        nonlocal index
        value = primary()
        while index < len(tokens) and tokens[index] == '|':
            index += 1; value |= primary()
        return value
    result = bit_or()
    if index != len(tokens): raise ValueError('unsupported trailing tokens')
    return result

def audit(text, census):
    macros, declarations, directives, typedefs, prototypes, unaccounted = scan(text)
    kinds = {'struct','struct-tag','pointer-typedef','scalar-typedef','function-pointer-typedef','function-prototype'}
    entries = census['entries']
    macro_entries = [e for e in entries if e['kind'] not in kinds]
    observed_macros = {(e['symbol'], e['line_start']) for e in macro_entries}
    observed_decls = {(e['kind'], e['symbol'], e['line_start'], e['line_end']) for e in entries if e['kind'] in kinds}
    issues = []
    if unaccounted: issues.append(f'unaccounted source lines: {unaccounted}')
    for label, diff in [('macro source-only', macros.keys()-observed_macros), ('macro census-only', observed_macros-macros.keys()),
                        ('declaration source-only', declarations-observed_decls), ('declaration census-only', observed_decls-declarations)]:
        if diff: issues.append(f'{label}: {sorted(diff)}')
    definitions = {name: expr for (name, line), expr in macros.items()}
    if len(definitions) != len(macros): issues.append('repeated macro names require conditional review')
    for entry in macro_entries:
        key = (entry['symbol'], entry['line_start'])
        if key not in macros: continue
        expression = macros[key]
        if re.sub(r'\s+', '', expression) != re.sub(r'\s+', '', entry.get('expression', '')):
            issues.append(f'source expression mismatch: {key}')
        try:
            value = independent_value(expression, definitions, (key[0],))
            if type(entry.get('value')) is not int or entry['value'] != value:
                issues.append(f'effective value mismatch: {key}')
        except ValueError as e:
            issues.append(f'unresolved macro {key}: {e}')
    return {'physical_lines':len(text.splitlines()), 'directives':dict(sorted(directives.items())),
            'raw_macro_names':len(macros), 'macro_values_compared':len(macro_entries), 'typedef_statements':typedefs, 'prototype_statements':prototypes,
            'source_declaration_names':len(declarations), 'unaccounted_lines':unaccounted, 'differences':issues}

def main():
    p = argparse.ArgumentParser(description=__doc__); p.add_argument('census'); args = p.parse_args()
    try:
        raw = subprocess.check_output(['git','show',f'{REV}:{PATH}'])
        with open(args.census, encoding='utf8') as f: c = json.load(f)
        result = audit(raw.decode('latin1'), c)
        print(json.dumps(result, indent=2))
        print('MECHANICAL CROSS-CHECK ONLY — independent review and completeness NOT ASSESSED')
        return 1 if result['differences'] else 0
    except (OSError, ValueError, KeyError, subprocess.CalledProcessError) as e:
        print(f'FAIL: {e}', file=sys.stderr); return 1

if __name__ == '__main__': sys.exit(main())
