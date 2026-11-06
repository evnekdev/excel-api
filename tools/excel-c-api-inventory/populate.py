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
DETAIL = {'expression', 'value', 'underlying', 'target', 'tag', 'signature', 'conditional_fields', 'linkage', 'guard', 'alias_of', 'subfamily'}

TYPE_DESCRIPTIONS = {
    'BOOL': 'Boolean scalar declared as INT32; external Windows type definitions determine its platform contract.',
    'XCHAR': 'Excel 12 character scalar declared as Windows WCHAR.',
    'RW': 'Excel 12 row scalar declared as INT32.',
    'COL': 'Excel 12 column scalar declared as INT32.',
    'IDSHEET': 'Excel sheet identifier declared as Windows DWORD_PTR; no fixed pointer width is asserted.',
    'XLREF': 'Legacy rectangular reference with WORD row bounds and BYTE column bounds.',
    'XLMREF': 'Legacy multi-area reference with a WORD count and a trailing XLREF table declared with one element.',
    'XLREF12': 'Excel 12 rectangular reference with RW row bounds and COL column bounds.',
    'XLMREF12': 'Excel 12 multi-area reference with a WORD count and a trailing XLREF12 table declared with one element.',
    'FP': 'Numeric array with unsigned-short row and column counts and a trailing double array declared with one element.',
    'FP12': 'Excel 12 numeric array with INT32 row and column counts and a trailing double array declared with one element.',
    'XLOPER': 'Legacy tagged Excel value union; the Boolean member is WORD xbool in C++ and WORD bool in C.',
    'XLOPER12': 'Excel 12 tagged value union with XCHAR strings, BOOL xbool, RW/COL array dimensions and a DWORD type field.',
    'FMLAINFO': 'Formula-edit information: mode, character count, read-only formula buffer, selection offsets and caret offset.',
    'MOUSEINFO': 'Mouse-query structure with HWND/POINT inputs and document type, hit-test code, row and column outputs.',
}
RETURNS = {
    'Success':'successful callback', 'Abort':'halted macro', 'InvXlfn':'invalid function number',
    'InvCount':'invalid argument count', 'InvXloper':'invalid OPER structure',
    'StackOvfl':'stack overflow', 'Failed':'failed command', 'Uncalced':'uncalculated cell',
    'NotThreadSafe':'call disallowed during multithreaded calculation',
    'InvAsynchronousContext':'invalid asynchronous function handle',
    'NotClusterSafe':'operation unsupported on a cluster',
}
HITS = {
    'None':'none of the listed hit areas', 'Client':'internal client area',
    'VSplit':'vertical pane split', 'HSplit':'horizontal pane split',
    'ColWidth':'column-width adjuster', 'RwHeight':'row-height adjuster',
    'RwColHdr':'intersection of row and column headers', 'Object':'object body',
    'TopLeft':'top-left object sizing handle', 'BotLeft':'bottom-left object sizing handle',
    'Left':'left object sizing handle', 'TopRight':'top-right object sizing handle',
    'BotRight':'bottom-right object sizing handle', 'Right':'right object sizing handle',
    'Top':'top object sizing handle', 'Bot':'bottom object sizing handle',
    'RwGut':'row outline gutter', 'ColGut':'column outline gutter',
    'TextBox':'text-box body', 'RwLevels':'row outline-level buttons',
    'ColLevels':'column outline-level buttons', 'Dman':'selection drag/drop handle',
    'DmanFill':'selection autofill handle', 'XSplit':'intersection of pane splits',
    'Vertex':'polygon vertex', 'AddVtx':'polygon vertex in add mode',
    'DelVtx':'polygon vertex in delete mode', 'RwHdr':'row header', 'ColHdr':'column header',
    # The source comments for these two are inconsistent; preserve uncertainty.
    'RwShow':'hidden-dimension reveal handle (source row/column wording is inconsistent)',
    'ColShow':'column counterpart of htRwShow (source wording requires behavioral review)',
    'Sizing':'internal sizing area', 'Sxpivot':'pivot-table drag/drop tile',
    'Tabs':'sheet paging tabs', 'Edit':'internal edit area',
}
SPECIAL = {
    'xlGetInst': 'Application instance identifier returned as an integer; the header limits this selector to 32-bit platforms.',
    'xlGetInstPtr': 'Application instance handle selector documented for both 32-bit and 64-bit platforms.',
    'xlGetFmlaInfo': 'Formula-edit information selector restricted by the header to LPenHelper calls.',
    'xlGetMouseInfo': 'Mouse information selector restricted by the header to LPenHelper calls.',
    'xlAsyncReturn': 'Auxiliary selector for returning an asynchronous function result.',
    'xlEventRegister': 'Auxiliary selector for registering an XLL event.',
    'xlRunningOnCluster': 'Auxiliary selector for querying whether execution is on a compute cluster.',
}

def description(c):
    """Brief source-grounded explanations; no inferred modern call legality."""
    name, kind = c['symbol'], c['kind']
    if name in TYPE_DESCRIPTIONS: return TYPE_DESCRIPTIONS[name]
    if kind == 'struct-tag': return f"C struct tag naming the aggregate typedef {c['target']}; not a second storage layout."
    if kind == 'pointer-typedef': return f"Named pointer to {c['target']} declared alongside that aggregate; no ownership transfer is implied."
    if kind == 'function-pointer-typedef':
        return 'Cluster connector asynchronous callback pointer returning int with CALLBACK convention and handle/result LPXLOPER12 parameters.'
    if kind == 'function-prototype':
        if name in ('Excel4', 'Excel12'):
            return f"Variadic _cdecl Excel callback using {'LPXLOPER12' if name == 'Excel12' else 'LPXLOPER'} values and an explicit argument count; returns int."
        if name in ('Excel4v', 'Excel12v'):
            return f"Array-argument pascal Excel callback using {'LPXLOPER12' if name == 'Excel12v' else 'LPXLOPER'} values and an explicit argument count; returns int."
        if name == 'XLCallVer': return 'No-argument pascal call-interface version prototype returning int; actual host results are unassessed.'
        return 'pascal helper prototype returning long and accepting an int selector plus VOID pointer; used by the formula and mouse information selectors.'
    family = c.get('subfamily')
    if family == 'cluster-return-code':
        meaning = {'xlHpcRetSuccess':'success', 'xlHpcRetSessionIdInvalid':'invalid session identifier', 'xlHpcRetCallFailed':'failed call'}[name]
        return f'Cluster connector entry-point result code for {meaning}; distinct from Excel callback xlret codes.'
    if family == 'edit-mode':
        return f"Formula edit-mode code for {name[6:].lower()} mode; xlModeReady denotes no active edit."
    if family == 'document-type':
        meaning = {'dtNil':'no qualifying document window or a nonselected window at idle', 'dtSheet':'worksheet',
                   'dtProc':'XLM macro sheet', 'dtChart':'chart', 'dtBasic':'VBA'}[name]
        return f'Document/page discriminator for {meaning}, associated with MOUSEINFO.dt.'
    if family == 'mouse-hit-test': return f'Mouse hit-test code for {HITS[name[2:]]}, associated with MOUSEINFO.ht.'
    if family == 'user-defined-function-id': return 'User-defined function dispatch number; the header specifies a function reference as the first argument.'
    if kind == 'return-code': return f"Excel4/Excel4v/Excel12/Excel12v return code indicating {RETURNS[name[5:]]}."
    if kind == 'error-code': return f"Excel error-value code {name[5:]} stored in val.err when the value is tagged as an error."
    if kind == 'flow-code': return f"Flow-control discriminator for {name[6:].lower()}, stored in val.flow.xlflow; not an xlf function number."
    if kind == 'event-id':
        return 'Event identifier supplied to xlEventRegister for ' + ('completed calculation.' if name.endswith('Ended') else 'interrupted calculation.')
    if kind == 'ownership-flag': return f"Ownership-related xltype bit {name}; separate from the value-type discriminator and requiring an independently verified free protocol."
    if kind == 'type-flag': return f"XLOPER/XLOPER12 xltype discriminator for {name[6:]} values" + (' formed from xltypeStr and xltypeInt.' if name == 'xltypeBigData' else '.')
    if kind == 'function-modifier': return f"Function-number bit {name} ({c['expression']}); distinct from the base function or command number."
    if kind == 'special-function-id':
        return SPECIAL.get(name, f"C API auxiliary selector {name}, combined with xlSpecial; header presence alone does not establish call legality.")
    if kind == 'worksheet-function-id': return f"SDK function number for {name[3:]}, in the combined Excel/XLM function table; modern availability and argument semantics are unassessed."
    if kind == 'macro-command-id': return f"SDK command number for {name[3:]}, combined with xlCommand; command context and host support require separate verification."
    raise ValueError(f'no reviewed description family for {kind}: {name}')

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
            old_template = f"{LABELS[c['kind']]} {c['symbol']} in the Excel 2013 XLL SDK."
            if r['description'] == old_template:
                r['description'] = description(c)
            continue
        if next_id > 999999: raise ValueError('CAPI ID space exhausted')
        ident = f'CAPI-{next_id:06d}'; next_id += 1
        r = {'id':ident, 'symbol':c['symbol'], 'kind':c['kind'],
             'description':description(c),
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
