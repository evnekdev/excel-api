"""Append source-derived catalogue candidates without changing existing IDs.

Use --write only after reviewing source-census and schema changes. Re-running
preserves all existing identities and assigns only previously unused IDs.
"""
import argparse
import json
from pathlib import Path
import source_census as source
import validate

LABELS = {
    'struct':'C aggregate typedef', 'struct-tag':'C struct tag',
    'scalar-typedef':'scalar typedef', 'pointer-typedef':'pointer typedef',
    'function-pointer-typedef':'callback function-pointer typedef',
    'function-prototype':'Excel callback function prototype',
    'type-flag':'Excel value-type flag', 'ownership-flag':'Excel memory ownership flag',
    'error-code':'Excel error code', 'flow-code':'Excel flow code',
    'return-code':'Excel callback return code', 'event-id':'Excel event identifier',
    'function-modifier':'Excel function-number modifier',
    'special-function-id':'Excel C API auxiliary callback identifier',
    'worksheet-function-id':'historical Excel/XLM function identifier',
    'macro-command-id':'Excel/XLM macro command identifier',
    'other-constant':'Excel SDK constant',
}
DETAIL = {'expression', 'value', 'underlying', 'target', 'tag', 'signature', 'conditional_fields', 'linkage', 'guard', 'alias_of'}

def build(records, candidates, history):
    by_identity = {(r['kind'], r['symbol']): r for r in records}
    if len(by_identity) != len(records): raise ValueError('duplicate catalogue kind/symbol')
    reserved = set(history) | {r['id'] for r in records}
    next_id = max(int(i[5:]) for i in reserved) + 1
    added = []
    for c in candidates:
        key = (c['kind'], c['symbol'])
        details = {k:c[k] for k in DETAIL if k in c}
        if key in by_identity:
            r = by_identity[key]
            s = r['source']
            if (s['line_start'], s['line_end'], s['revision'], s['blob_sha']) != (c['line_start'], c['line_end'], source.REVISION, source.BLOB):
                raise ValueError(f'existing source identity/anchor differs: {key}')
            for k in DETAIL:
                if k in r and (k not in details or r[k] != details[k]):
                    raise ValueError(f'existing source detail differs: {key} {k}')
            r['schema_version'] = 2
            r.update(details)
            continue
        if next_id > 999999: raise ValueError('CAPI ID space exhausted')
        ident = f'CAPI-{next_id:06d}'; next_id += 1
        r = {'id':ident, 'symbol':c['symbol'], 'kind':c['kind'],
             'description':f"{LABELS[c['kind']]} {c['symbol']} in the Excel 2013 XLL SDK.",
             'source':{'repository':'evnekdev/excel-api','revision':source.REVISION,'path':source.PATH,'blob_sha':source.BLOB,
                       'line_start':c['line_start'],'line_end':c['line_end']},
             'edna_ids':[], 'rust_status':'unassessed', 'rust_evidence':[], 'excel_host_status':'not-run',
             'excel_host_evidence':[], 'behavior_refs':[], 'schema_version':2, **details}
        added.append(r)
    found = {(c['kind'], c['symbol']) for c in candidates}
    missing = set(by_identity) - found
    if missing: raise ValueError(f'existing catalogue identity absent from source candidates: {missing}')
    return records + added

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--write', action='store_true')
    args = parser.parse_args()
    records = validate.read_jsonl(validate.CATALOGUE.read_bytes())
    history = validate.historical_identities()
    validate.validate(records, history)
    rows = build(records, source.extract(source.pinned_header()), history)
    validate.validate(rows, history)
    payload = ''.join(json.dumps(r, ensure_ascii=False, separators=(',', ':')) + '\n' for r in rows)
    if args.write:
        validate.CATALOGUE.write_text(payload, encoding='utf8', newline='\n')
        print(f'wrote {len(rows)} records, including {len(rows)-len(records)} new IDs')
    else:
        print(payload, end='')

if __name__ == '__main__': main()
