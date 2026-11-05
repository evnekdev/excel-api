#!/usr/bin/env python3
"""Scoped validator for the version-1 Excel C API catalogue; see README.md."""
import argparse
import json
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
CATALOGUE = ROOT / 'knowledge/excel-c-api/symbols.jsonl'
BASELINE = '172a003172e5d90a26ea4a8345218988d4b189cc'
CATALOGUE_PATH = 'knowledge/excel-c-api/symbols.jsonl'
SOURCE_PATH = 'tools/Excel2013XLLSDK/INCLUDE/XLCALL.H'
KINDS = {'struct', 'type-flag', 'special-function-id', 'worksheet-function-id', 'macro-command-id', 'function-prototype'}
RUST = {'unassessed', 'absent', 'partial', 'implemented', 'not-applicable'}
HOST = {'not-run', 'pass', 'fail', 'blocked'}
SHA = re.compile(r'[0-9a-f]{40}\Z')
NAME = re.compile(r'[A-Za-z_][A-Za-z_0-9]*\Z')

class Invalid(ValueError):
    pass

class HistoryUnavailable(ValueError):
    pass

def git(*args):
    p = subprocess.run(['git', *args], cwd=ROOT, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if p.returncode:
        raise Invalid('git ' + ' '.join(args) + ': ' + p.stderr.decode('utf-8', 'replace').strip())
    return p.stdout

def read_jsonl(raw):
    try:
        text = raw.decode('utf-8')
        records = []
        for n, line in enumerate(text.splitlines(), 1):
            if not line.strip():
                continue
            try:
                r = json.loads(line, object_pairs_hook=unique_keys, parse_constant=lambda x: (_ for _ in ()).throw(ValueError('non-finite number')))
            except (ValueError, TypeError) as e:
                raise Invalid(f'line {n}: invalid JSON: {e}') from e
            if not isinstance(r, dict):
                raise Invalid(f'line {n}: expected object')
            records.append(r)
        if not records:
            raise Invalid('empty catalogue')
        return records
    except UnicodeDecodeError as e:
        raise Invalid(f'invalid UTF-8: {e}') from e

def unique_keys(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f'duplicate key {key}')
        result[key] = value
    return result

def identity(r):
    return (r['source']['path'], r['kind'], r['symbol'])

def historical_identities(baseline=BASELINE, head='HEAD', upstream='refs/remotes/origin/master', extra=None):
    """Reserve identities from catalogue snapshots reachable from both relevant tips."""
    try:
        if git('rev-parse', '--is-shallow-repository').decode().strip() != 'false':
            raise HistoryUnavailable('shallow repository; fetch complete catalogue history')
        tips = [head, upstream] + ([extra] if extra else [])
        for ref in (baseline, *tips):
            git('rev-parse', '--verify', f'{ref}^{{commit}}')
        for ref in tips:
            git('merge-base', '--is-ancestor', baseline, ref)
        commits = git('log', '--full-history', '--format=%H', *tips, '--', CATALOGUE_PATH).decode().splitlines()
        if git('rev-parse', f'{baseline}^{{commit}}').decode().strip() not in commits:
            raise HistoryUnavailable('original catalogue baseline is missing from relevant history')
    except Invalid as e:
        raise HistoryUnavailable(f'required catalogue history unavailable: {e}') from e
    reserved = {}; snapshots = {}
    def snapshot(commit):
        if commit in snapshots:
            return snapshots[commit]
        try:
            if not git('ls-tree', commit, '--', CATALOGUE_PATH).strip():
                snapshots[commit] = None
                return None
            records = read_jsonl(git('show', f'{commit}:{CATALOGUE_PATH}'))
        except Invalid as e:
            raise HistoryUnavailable(f'cannot read catalogue snapshot {commit}: {e}') from e
        seen = {}
        for r in records:
            try:
                ident = r['id']; key = identity(r)
                if not isinstance(ident, str) or not re.fullmatch(r'CAPI-[0-9]{6}', ident) or ident == 'CAPI-000000':
                    raise Invalid(f'invalid historical ID in {commit}')
                if any(not isinstance(part, str) or not part for part in key):
                    raise Invalid(f'invalid historical identity for {ident} in {commit}')
                if ident in seen:
                    raise Invalid(f'duplicate historical ID {ident} in {commit}')
                seen[ident] = key
                if ident in reserved and reserved[ident] != key:
                    raise Invalid(f'historical ID {ident} reassigned in reachable catalogue history')
                reserved[ident] = key
            except (KeyError, TypeError) as e:
                raise Invalid(f'invalid historical record in {commit}: {e}') from e
        snapshots[commit] = seen
        return seen
    for commit in commits:
        current = snapshot(commit)
        if current is None:
            raise Invalid(f'historical catalogue removed in {commit}; no retirement policy exists')
        try:
            parents = git('rev-list', '--parents', '-n', '1', commit).decode().split()[1:]
        except Invalid as e:
            raise HistoryUnavailable(f'cannot read parents of {commit}: {e}') from e
        for parent in parents:
            previous = snapshot(parent)
            if previous:
                missing = previous.keys() - current.keys()
                if missing:
                    raise Invalid(f'historical IDs removed in {commit}: ' + ', '.join(sorted(missing)))
    return reserved

def strip_comments(text):
    # Preserve newlines so anchors remain meaningful. Strings are irrelevant to
    # the narrow supported declaration forms, and are rejected in macros.
    return re.sub(r'/\*.*?\*/|//[^\n]*', lambda m: re.sub(r'[^\n]', ' ', m.group()), text, flags=re.S)

def source_blob(source):
    if source['repository'] != 'evnekdev/excel-api' or source['path'] != SOURCE_PATH:
        raise Invalid('unsupported source repository or path')
    for k in ('revision', 'blob_sha'):
        if not isinstance(source[k], str) or not SHA.fullmatch(source[k]):
            raise Invalid(f'invalid source {k}')
    rev, path = source['revision'], source['path']
    if git('rev-parse', f'{rev}^{{commit}}').decode().strip() != rev:
        raise Invalid('source revision is not an exact commit')
    actual = git('rev-parse', f'{rev}:{path}').decode().strip()
    if actual != source['blob_sha']:
        raise Invalid('source blob does not match revision and path')
    return git('cat-file', 'blob', actual).decode('latin-1')

def declaration(r, text):
    s = r['source']; lines = strip_comments(text).splitlines()
    a, b = s['line_start'], s['line_end']
    if type(a) is not int or type(b) is not int or not (1 <= a <= b <= len(lines)):
        raise Invalid('invalid or out-of-range source lines')
    span = '\n'.join(lines[a-1:b]); name = re.escape(r['symbol']); kind = r['kind']
    if kind == 'struct':
        match = re.search(r'\btypedef\s+struct\b[^{;]*\{', span)
        if not match:
            raise Invalid('anchor is not a typedef struct')
        depth = 1; i = match.end()
        while i < len(span) and depth:
            depth += (span[i] == '{') - (span[i] == '}'); i += 1
        if depth or not re.match(r'\s*' + name + r'\s*[,;]', span[i:]):
            raise Invalid('struct typedef name or braces do not match')
    elif kind == 'function-prototype':
        if not re.search(r'(?m)^\s*(?:int|long|void)\s+(?:_cdecl|pascal|PASCAL)\s+' + name + r'\s*\([^;{}]*\)\s*;', span):
            raise Invalid('anchor is not the supported function prototype')
    else:
        m = re.search(r'(?m)^\s*#\s*define\s+' + name + r'\s+([^\n]+)$', span)
        if not m:
            raise Invalid('anchor is not the named object-like macro')
        expr = m.group(1).strip()
        if not re.fullmatch(r'[A-Za-z_0-9xXa-fA-F()|+\-\s]+', expr):
            raise Invalid('unsupported macro expression; needs review')
    return True

def validate(records, baseline=None, blob_reader=source_blob):
    ids = set(); identities = set(); cache = {}
    for r in records:
        try:
            allowed = {'schema_version', 'id', 'symbol', 'kind', 'description', 'source', 'edna_ids', 'rust_status', 'rust_evidence', 'excel_host_status', 'excel_host_evidence', 'behavior_refs'}
            if set(r) - allowed:
                raise Invalid('unsupported fields (including aliases or guards): ' + ', '.join(sorted(set(r) - allowed)))
            for k in ('id', 'symbol', 'kind', 'description', 'rust_status', 'excel_host_status'):
                if not isinstance(r.get(k), str) or not r[k].strip():
                    raise Invalid(f'missing or invalid {k}')
            if not re.fullmatch(r'CAPI-[0-9]{6}', r['id']) or r['id'] == 'CAPI-000000':
                raise Invalid('invalid CAPI ID')
            if not NAME.fullmatch(r['symbol']) or r['kind'] not in KINDS:
                raise Invalid('invalid symbol or kind')
            if r.get('schema_version', 1) != 1 or type(r.get('schema_version', 1)) is not int:
                raise Invalid('unsupported schema version')
            for k in ('edna_ids', 'rust_evidence', 'excel_host_evidence', 'behavior_refs'):
                if not isinstance(r.get(k), list) or any(not isinstance(x, str) or not x.strip() for x in r[k]):
                    raise Invalid(f'missing or invalid {k}')
            if len(set(r['edna_ids'])) != len(r['edna_ids']) or any(not re.fullmatch(r'EDNA-(?:0[1-9]|1[0-9]|2[0-5])', x) for x in r['edna_ids']):
                raise Invalid('invalid or duplicate EDNA reference')
            if r['rust_status'] not in RUST or r['excel_host_status'] not in HOST:
                raise Invalid('invalid status')
            for status, evidence, unassessed in [('rust_status', 'rust_evidence', 'unassessed'), ('excel_host_status', 'excel_host_evidence', 'not-run')]:
                if bool(r[evidence]) != (r[status] != unassessed):
                    raise Invalid(f'{status} requires appropriate {evidence}')
            if r['excel_host_status'] == 'pass':
                raise Invalid('host pass cannot be verified from schema v1 evidence strings; needs reviewed structured real-Excel evidence')
            s = r.get('source')
            if not isinstance(s, dict) or any(k not in s for k in ('repository', 'revision', 'path', 'blob_sha', 'line_start', 'line_end')):
                raise Invalid('missing source fields')
            if set(s) != {'repository', 'revision', 'path', 'blob_sha', 'line_start', 'line_end'}:
                raise Invalid('unsupported source fields; aliases and conditional variants need schema review')
            key = identity(r)
            if r['id'] in ids: raise Invalid('duplicate ID')
            if key in identities: raise Invalid('duplicate canonical identity')
            ids.add(r['id']); identities.add(key)
            if baseline and r['id'] in baseline and baseline[r['id']] != key:
                raise Invalid('historical ID reassigned')
            cachekey = (s['repository'], s['revision'], s['path'], s['blob_sha'])
            if cachekey not in cache: cache[cachekey] = blob_reader(s)
            declaration(r, cache[cachekey])
        except (Invalid, KeyError, TypeError, UnicodeDecodeError) as e:
            raise Invalid(f"{r.get('id', '<unknown>')}: {e}") from e
    if baseline is not None:
        missing = baseline.keys() - ids
        if missing:
            raise Invalid('historical IDs removed without a retirement policy: ' + ', '.join(sorted(missing)))
    return len(records)

def compare_census(records, census):
    if census is None:
        return ['NOT ASSESSED — complete source census pending #117'], False
    if not isinstance(census, dict) or not isinstance(census.get('entries'), list) or type(census.get('complete')) is not bool:
        raise Invalid('invalid census: expected complete boolean and entries array')
    found = {}; messages = []; problems = False
    for entry in census['entries']:
        if not isinstance(entry, dict) or any(not isinstance(entry.get(k), str) or not entry[k] for k in ('path', 'kind', 'symbol', 'disposition')):
            raise Invalid('invalid census entry')
        key = (entry['path'], entry['kind'], entry['symbol'])
        if key in found:
            messages.append(f'DUPLICATE source identity: {key}'); problems = True
        found[key] = entry
        if entry['disposition'] not in ('eligible', 'excluded', 'ambiguous'):
            raise Invalid('invalid census disposition')
        if entry['disposition'] != 'eligible' and not entry.get('reason'):
            raise Invalid('excluded/ambiguous entry needs reason')
        if entry['disposition'] != 'eligible':
            messages.append(f"{entry['disposition'].upper()}: {key}: {entry['reason']}")
        if entry['disposition'] == 'ambiguous': problems = True
    actual = {identity(r): r for r in records}
    for key, entry in found.items():
        if entry['disposition'] == 'eligible' and key not in actual:
            messages.append(f'MISSING: {key}'); problems = True
        if entry['disposition'] == 'excluded' and key in actual:
            messages.append(f'EXCLUDED BUT CATALOGUED: {key}'); problems = True
        if entry.get('alias_of'): messages.append(f"ALIAS: {key} -> {entry['alias_of']}")
        if entry.get('guard'): messages.append(f"CONDITIONAL: {key}: {entry['guard']}")
        if key in actual and entry.get('source_blob') and entry['source_blob'] != actual[key]['source']['blob_sha']:
            messages.append(f'CHANGED SOURCE BLOB: {key}'); problems = True
        if key in actual and 'expression' in entry:
            r = actual[key]
            if not isinstance(entry['expression'], str) or r['kind'] in ('struct', 'function-prototype'):
                raise Invalid('expression is supported only for macro census entries')
            s = r['source']; lines = strip_comments(source_blob(s)).splitlines()
            span = '\n'.join(lines[s['line_start']-1:s['line_end']])
            match = re.search(r'(?m)^\s*#\s*define\s+' + re.escape(r['symbol']) + r'\s+([^\n]+)$', span)
            if not match or re.sub(r'\s+', '', match.group(1)) != re.sub(r'\s+', '', entry['expression']):
                messages.append(f'CHANGED MACRO EXPRESSION: {key}'); problems = True
    if census['complete']:
        for key in actual.keys() - found.keys():
            messages.append(f'EXTRA: {key}'); problems = True
    if not census['complete']:
        messages.append('NOT ASSESSED — complete source census pending #117')
    return messages, problems

def freshness(records):
    out = ['Source evidence: pinned historical blobs; integrity checked by CLI.',
           'Catalogue completeness: NOT ASSESSED — complete source census pending #117']
    for field, status, label in [('rust_status', 'unassessed', 'Rust implementation'), ('excel_host_status', 'not-run', 'Native Excel evidence')]:
        assessed = sum(r[field] != status for r in records)
        out.append(f'{label}: ' + ('NOT ASSESSED' if not assessed else f'UNKNOWN / insufficient structured revision and dependency evidence ({assessed} assessed records)'))
    out.append('Higher-level EDNA functional coverage: NOT ASSESSED')
    return out

def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('command', choices=['check', 'reconcile', 'freshness'])
    parser.add_argument('--catalogue', type=Path, default=CATALOGUE)
    parser.add_argument('--baseline', help='additional Git history tip to protect; does not replace required history')
    parser.add_argument('--census', type=Path, help='independently reviewed census JSON')
    args = parser.parse_args(argv)
    try:
        records = read_jsonl(args.catalogue.read_bytes())
        # Validate the input even if the required Git history is unavailable.
        validate(records)
        try:
            history = historical_identities(extra=args.baseline)
        except HistoryUnavailable as e:
            print(f'NOT VERIFIED: historical ID protection incomplete: {e}', file=sys.stderr)
            return 3
        n = validate(records, history)
        if args.command == 'check':
            print(f'PASS: {n} records, schema v1 and pinned declaration anchors; reachable catalogue history checked')
            print('LIMIT: macro values/aliases/guards, full C semantics and evidence authenticity are not verified.')
        elif args.command == 'reconcile':
            census = json.loads(args.census.read_text(encoding='utf-8'), object_pairs_hook=unique_keys) if args.census else None
            messages, problems = compare_census(records, census)
            print('\n'.join(messages))
            if problems:
                print('RECONCILIATION MISMATCH — SOURCE COMPLETENESS NOT VERIFIED')
                return 2
            if census is not None:
                scope = '' if census['complete'] else 'PARTIAL '
                print(f'MATCHES SUPPLIED {scope}CENSUS — SOURCE COMPLETENESS NOT VERIFIED')
            return 3  # A caller-supplied census cannot prove source completeness.
        else:
            print('\n'.join(freshness(records)))
        return 0
    except (Invalid, OSError, ValueError) as e:
        print(f'FAIL: {e}', file=sys.stderr); return 1

if __name__ == '__main__':
    sys.exit(main())
