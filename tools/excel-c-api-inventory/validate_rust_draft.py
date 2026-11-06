"""Validate the preliminary #122 candidate inventory, not implementation coverage."""
import json
import pathlib
import re
import subprocess
import sys

ROOT = pathlib.Path(__file__).resolve().parents[2]
REVISION = '172b306fae9b8496d5ad79b626a9fa0b1890cfd4'
CATALOGUE = 'knowledge/excel-c-api/symbols.jsonl'
DRAFT = ROOT / 'knowledge/excel-c-api/rust-assessment-draft.jsonl'

def git_show(path):
    return subprocess.check_output(['git', 'show', f'{REVISION}:{path}'], cwd=ROOT).decode('utf-8')

def verify(rows):
    catalogue = [json.loads(line) for line in git_show(CATALOGUE).splitlines()]
    ids = {row['id']: row for row in catalogue}
    if len(ids) != len(catalogue): raise ValueError('duplicate source catalogue ID')
    found = set()
    allowed = {'crates/excel-api-sys/src/constants.rs', 'crates/excel-api-sys/src/types.rs', 'crates/excel-api-sys/src/functions.rs'}
    cache = {}
    for row in rows:
        if set(row) != {'capi_id','raw_state','functional_state','test_state','host_state','exact_name_candidates'}: raise ValueError('invalid row fields')
        ident = row['capi_id']
        if ident not in ids or ident in found: raise ValueError(f'duplicate or invalid ID: {ident}')
        found.add(ident)
        if any(row[k] != 'unassessed' for k in ('raw_state','functional_state','test_state','host_state')): raise ValueError(f'unsupported assessment: {ident}')
        if not isinstance(row['exact_name_candidates'],list): raise ValueError(f'invalid candidate list: {ident}')
        candidate_keys = set()
        for candidate in row['exact_name_candidates']:
            if set(candidate) != {'path','line','symbol'} or candidate['path'] not in allowed or candidate['symbol'] != ids[ident]['symbol'] or type(candidate['line']) != int: raise ValueError(f'invalid candidate: {ident}')
            key = (candidate['path'], candidate['line'], candidate['symbol'])
            if key in candidate_keys: raise ValueError(f'duplicate candidate: {ident}')
            candidate_keys.add(key)
            path = candidate['path']
            if path not in cache: cache[path] = git_show(path).splitlines()
            line = candidate['line']
            if not 1 <= line <= len(cache[path]) or not re.match(r'\s*pub\s+(?:const|struct|union|enum|type)\s+' + re.escape(candidate['symbol']) + r'\b', cache[path][line-1]): raise ValueError(f'stale candidate: {ident}')
    if found != set(ids): raise ValueError(f'missing IDs: {len(set(ids)-found)}')
    return len(found)

def main():
    try:
        rows = [json.loads(line) for line in DRAFT.read_text(encoding='utf-8').splitlines()]
        print(f'PASS: {verify(rows)} IDs accounted for as unassessed; candidates resolve at {REVISION}')
        print('LIMIT: exact-name candidates are search leads, not verified representations or implementation coverage')
    except (ValueError, OSError, subprocess.CalledProcessError) as exc:
        print(f'FAIL: {exc}', file=sys.stderr); return 1
    return 0

if __name__ == '__main__': sys.exit(main())
