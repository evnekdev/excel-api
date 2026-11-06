"""Validate the preliminary #122 candidate inventory, not implementation coverage."""
import ast
import functools
import json
import pathlib
import re
import subprocess
import sys

ROOT = pathlib.Path(__file__).resolve().parents[2]
REVISION = '172b306fae9b8496d5ad79b626a9fa0b1890cfd4'
CATALOGUE = 'knowledge/excel-c-api/symbols.jsonl'
DRAFT = ROOT / 'knowledge/excel-c-api/rust-assessment-draft.jsonl'
ALTERNATE_NAMES = {
    'Excel12': 'Excel12Fn', 'Excel12v': 'Excel12vFn', 'XLCallVer': 'XLCallVerFn',
    'xlref12': 'XLREF12', 'xlmref12': 'XLMREF12', '_FP12': 'FP12',
    'xloper12': 'XLOPER12',
}
TEST_EVIDENCE = {
    'xlAbort': ('crates/excel-api/src/excel_call.rs', 'abort_descriptor_and_zero_retain_clear_forms_are_exact'),
    'xlCoerce': ('crates/excel-api/src/excel_call.rs', 'selected_m11_descriptors_preserve_sdk_ids_and_contracts'),
    'xlSheetId': ('crates/excel-api/src/excel_call.rs', 'selected_m11_descriptors_preserve_sdk_ids_and_contracts'),
    'xlSheetNm': ('crates/excel-api/src/excel_call.rs', 'selected_m11_descriptors_preserve_sdk_ids_and_contracts'),
    'xlfCaller': ('crates/excel-api/src/excel_call.rs', 'selected_m11_descriptors_preserve_sdk_ids_and_contracts'),
    'xltypeFlow': ('crates/excel-api/src/borrowed.rs', 'every_sdk_type_is_decoded_or_explicitly_rejected'),
    'xltypeBigData': ('crates/excel-api/src/borrowed.rs', 'every_sdk_type_is_decoded_or_explicitly_rejected'),
    'BOOL': ('crates/excel-api-sys/src/types.rs', 'fixed_width_primitives_match_the_sdk'),
    'XCHAR': ('crates/excel-api-sys/src/types.rs', 'fixed_width_primitives_match_the_sdk'),
    'RW': ('crates/excel-api-sys/src/types.rs', 'fixed_width_primitives_match_the_sdk'),
    'COL': ('crates/excel-api-sys/src/types.rs', 'fixed_width_primitives_match_the_sdk'),
    'IDSHEET': ('crates/excel-api-sys/src/types.rs', 'fixed_width_primitives_match_the_sdk'),
    'XLMREF12': ('crates/excel-api-sys/src/types.rs', 'flexible_array_headers_preserve_one_c_element'),
    'FP12': ('crates/excel-api-sys/src/types.rs', 'flexible_array_headers_preserve_one_c_element'),
    'XLOPER12': ('crates/excel-api-sys/src/types.rs', 'xloper12_has_windows_x64_sdk_layout'),
    'xlAsyncReturn': ('crates/excel-api/src/async_udf.rs', 'successful_completion_is_at_most_once_and_preserves_handle'),
    **{name: ('crates/excel-api/src/borrowed.rs', 'decodes_scalar_tags_and_masks_ownership_bits') for name in (
        'xltypeNum', 'xltypeBool', 'xltypeInt', 'xltypeErr', 'xlbitXLFree', 'xlbitDLLFree',
        'xlerrNull', 'xlerrDiv0', 'xlerrValue', 'xlerrRef', 'xlerrName', 'xlerrNum', 'xlerrNA', 'xlerrGettingData')},
    'xltypeMissing': ('crates/excel-api/src/borrowed.rs', 'missing_and_nil_are_independent'),
    'xltypeNil': ('crates/excel-api/src/borrowed.rs', 'missing_and_nil_are_independent'),
}
ERROR_VARIANTS = dict(zip(
    ('xlerrNull','xlerrDiv0','xlerrValue','xlerrRef','xlerrName','xlerrNum','xlerrNA','xlerrGettingData'),
    ('Null','Div0','Value','Ref','Name','Num','Na','GettingData')))

@functools.lru_cache(maxsize=None)
def git_show(path):
    return subprocess.check_output(['git', 'show', f'{REVISION}:{path}'], cwd=ROOT).decode('utf-8')

def numeric_constants():
    """Read the supported literal/OR subset of the pinned sys declarations."""
    values = {}
    def evaluate(node):
        if isinstance(node, ast.Constant) and type(node.value) is int:
            return node.value
        if isinstance(node, ast.Name) and node.id in values:
            return values[node.id]
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, ast.USub):
            return -evaluate(node.operand)
        if isinstance(node, ast.BinOp) and isinstance(node.op, ast.BitOr):
            return evaluate(node.left) | evaluate(node.right)
        raise ValueError('unsupported constant expression')
    for line in git_show('crates/excel-api-sys/src/constants.rs').splitlines():
        match = re.fullmatch(r'pub const (\w+): (?:u32|i32) = (.*);', line)
        if match:
            try:
                values[match.group(1)] = evaluate(ast.parse(match.group(2), mode='eval').body)
            except (ValueError, SyntaxError):
                pass
    return values

def public_definitions():
    paths = subprocess.check_output(
        ['git', 'ls-tree', '-r', '--name-only', REVISION, '--',
         'crates/excel-api-sys/src', 'crates/excel-api/src', 'crates/excel-api-macros/src'],
        cwd=ROOT, text=True, encoding='utf8').splitlines()
    definitions = {}
    for path in paths:
        if not path.endswith('.rs'):
            continue
        for number, line in enumerate(git_show(path).splitlines(), 1):
            match = re.match(r'\s*pub\s+(?:unsafe\s+)?(?:const|struct|union|enum|type|fn)\s+(\w+)', line)
            if match:
                definitions.setdefault(match.group(1), []).append({'path': path, 'line': number, 'symbol': match.group(1)})
    return definitions

def verify(rows):
    catalogue = [json.loads(line) for line in git_show(CATALOGUE).splitlines()]
    ids = {row['id']: row for row in catalogue}
    if len(ids) != len(catalogue): raise ValueError('duplicate source catalogue ID')
    found = set()
    constants = numeric_constants()
    definitions = public_definitions()
    allowed = {'crates/excel-api-sys/src/constants.rs', 'crates/excel-api-sys/src/types.rs', 'crates/excel-api-sys/src/functions.rs'}
    cache = {}
    for row in rows:
        if set(row) != {'capi_id','raw_state','functional_state','test_state','host_state','exact_name_candidates','functional_evidence','related_rust','test_evidence'}: raise ValueError('invalid row fields')
        ident = row['capi_id']
        if ident not in ids or ident in found: raise ValueError(f'duplicate or invalid ID: {ident}')
        found.add(ident)
        related = row['related_rust']
        if related is not None:
            if not isinstance(related, dict) or set(related) != {'path', 'line', 'symbol'} or related['symbol'] != ALTERNATE_NAMES.get(ids[ident]['symbol']) or related['path'] not in {'crates/excel-api-sys/src/types.rs', 'crates/excel-api-sys/src/functions.rs'} or type(related['line']) != int: raise ValueError(f'invalid related declaration: {ident}')
            related_lines = git_show(related['path']).splitlines()
            number = related['line']
            if not 1 <= number <= len(related_lines) or not re.match(r'\s*pub\s+(?:type|struct)\s+' + re.escape(related['symbol']) + r'\b', related_lines[number-1]): raise ValueError(f'stale related declaration: {ident}')
        if row['host_state'] != 'unassessed': raise ValueError(f'unsupported host assessment: {ident}')
        test = row['test_evidence']
        if row['test_state'] in ('test-source-present', 'passed-this-run'):
            symbol = ids[ident]['symbol']
            if not isinstance(test, dict) or set(test) != {'path','line','symbol','configuration','run'} or (test['path'],test['symbol']) != TEST_EVIDENCE.get(symbol) or type(test['line']) != int: raise ValueError(f'invalid test evidence: {ident}')
            expected_run = ('cargo test -p excel-api-sys --lib' if '/excel-api-sys/' in test['path'] else 'cargo test -p excel-api --lib')
            if test['run'] != (expected_run if row['test_state'] == 'passed-this-run' else None): raise ValueError(f'invalid test run: {ident}')
            lines = git_show(test['path']).splitlines(); number = test['line']
            if not 2 <= number <= len(lines) or lines[number-2].strip() != '#[test]' or lines[number-1].strip() != f'fn {test["symbol"]}() {{': raise ValueError(f'stale test evidence: {ident}')
            windows_x64 = '#[cfg(all(target_os = "windows", target_arch = "x86_64"))]' in lines[max(0, number-5):number-1]
            if test['configuration'] != ('windows-x86_64' if windows_x64 else 'default'): raise ValueError(f'invalid test configuration: {ident}')
            block = '\n'.join(lines[number-1:number+80]).split('\n    }', 1)[0]
            if not re.search(r'\b' + re.escape(symbol) + r'\b', block): raise ValueError(f'unrelated test: {ident}')
        elif row['test_state'] != 'unassessed' or test is not None: raise ValueError(f'unsupported test assessment: {ident}')
        if row['functional_state'] not in ('unassessed', 'descriptor-present', 'decode-branch-present', 'decoding-explicitly-unsupported', 'async-return-operation-present', 'ownership-mask-present', 'error-conversion-present'): raise ValueError(f'unsupported functional assessment: {ident}')
        if row['functional_state'] == 'descriptor-present':
            evidence = row['functional_evidence']
            if not isinstance(evidence, dict) or set(evidence) != {'path', 'symbol', 'line', 'configuration'} or evidence['path'] != 'crates/excel-api/src/excel_call.rs' or type(evidence['line']) != int: raise ValueError(f'invalid functional evidence: {ident}')
            source = git_show(evidence['path']).splitlines()
            number = evidence['line']
            if not 1 <= number <= len(source) or source[number - 1] != f"pub const {evidence['symbol']}: ExcelCallDescriptor = ExcelCallDescriptor {{": raise ValueError(f'stale functional evidence: {ident}')
            feature = '#[cfg(feature = "xlcontime-research")]' in source[max(0, number - 5):number - 1]
            if evidence['configuration'] != ('xlcontime-research' if feature else 'default'): raise ValueError(f'invalid descriptor configuration: {ident}')
            block = '\n'.join(source[number - 1:number + 25]).split('};', 1)[0]
            if f'name: "{ids[ident]["symbol"]}"' not in block or f'function: excel_api_sys::{ids[ident]["symbol"]}' not in block: raise ValueError(f'unrelated descriptor: {ident}')
        elif row['functional_state'] == 'error-conversion-present':
            evidence = row['functional_evidence']; symbol = ids[ident]['symbol']; variant = ERROR_VARIANTS.get(symbol)
            if not variant or not isinstance(evidence, dict) or set(evidence) != {'path','line','symbol','configuration'} or evidence['path'] != 'crates/excel-api/src/borrowed.rs' or evidence['symbol'] != variant or evidence['configuration'] != 'default' or type(evidence['line']) != int: raise ValueError(f'invalid error evidence: {ident}')
            lines = git_show(evidence['path']).splitlines(); n = evidence['line']
            if not 1 <= n <= len(lines) or lines[n-1].strip() != f'excel_api_sys::{symbol} => XlError::{variant},': raise ValueError(f'stale error evidence: {ident}')
            if f'{variant} = crate::{symbol},' not in (line.strip() for line in git_show('crates/excel-api-sys/src/errors.rs').splitlines()): raise ValueError(f'wrong raw error: {ident}')
            if f'excel_api_sys::XlError::{variant} => Self::{variant},' not in (line.strip() for line in git_show('crates/excel-api/src/error.rs').splitlines()): raise ValueError(f'wrong safe error: {ident}')
        elif row['functional_state'] == 'ownership-mask-present':
            evidence = row['functional_evidence']; symbol = ids[ident]['symbol']
            aliases = {'xlbitXLFree': 'XLBIT_XL_FREE', 'xlbitDLLFree': 'XLBIT_DLL_FREE'}
            if not isinstance(evidence, dict) or set(evidence) != {'path','line','symbol','configuration'} or evidence['path'] != 'crates/excel-api/src/borrowed.rs' or evidence['symbol'] != aliases.get(symbol) or evidence['configuration'] != 'default' or type(evidence['line']) != int: raise ValueError(f'invalid ownership evidence: {ident}')
            if f'pub const {evidence["symbol"]}: u32 = {symbol};' not in git_show('crates/excel-api-sys/src/constants.rs').splitlines(): raise ValueError(f'invalid ownership alias: {ident}')
            lines = git_show(evidence['path']).splitlines(); n = evidence['line']
            if not 1 <= n <= len(lines) or lines[n-1].strip() != 'let base_type = xltype & !(XLBIT_XL_FREE | XLBIT_DLL_FREE);': raise ValueError(f'stale ownership evidence: {ident}')
        elif row['functional_state'] == 'async-return-operation-present':
            evidence = row['functional_evidence']
            if ids[ident]['symbol'] != 'xlAsyncReturn' or not isinstance(evidence, dict) or set(evidence) != {'path','line','symbol','configuration'} or evidence['path'] != 'crates/excel-api/src/async_udf.rs' or evidence['symbol'] != 'return_to_excel' or evidence['configuration'] != 'default' or type(evidence['line']) != int: raise ValueError(f'invalid async evidence: {ident}')
            lines = git_show(evidence['path']).splitlines(); n = evidence['line']
            if not 1 <= n <= len(lines) or lines[n-1].strip() != 'fn return_to_excel(' or 'excel_api_sys::xlAsyncReturn' not in '\n'.join(lines[n-1:n+45]): raise ValueError(f'stale async evidence: {ident}')
        elif row['functional_state'] in ('decode-branch-present', 'decoding-explicitly-unsupported'):
            evidence = row['functional_evidence']; symbol = ids[ident]['symbol']
            if not isinstance(evidence, dict) or set(evidence) != {'path','line','symbol','configuration'} or evidence['path'] != 'crates/excel-api/src/borrowed.rs' or evidence['configuration'] != 'default' or type(evidence['line']) != int: raise ValueError(f'invalid decode evidence: {ident}')
            alias = evidence['symbol']
            if f'pub const {alias}: u32 = {symbol};' not in git_show('crates/excel-api-sys/src/constants.rs').splitlines(): raise ValueError(f'invalid decode alias: {ident}')
            lines = git_show(evidence['path']).splitlines(); n = evidence['line']
            if not 1 <= n <= len(lines) or not lines[n-1].strip().startswith(f'{alias} =>'): raise ValueError(f'stale decode evidence: {ident}')
            unsupported = 'Err(DecodeError::UnsupportedType' in lines[n-1]
            if unsupported != (row['functional_state'] == 'decoding-explicitly-unsupported'): raise ValueError(f'invalid decode claim: {ident}')
        elif row['functional_evidence'] is not None: raise ValueError(f'unassessed descriptor contains evidence: {ident}')
        if row['raw_state'] not in ('unassessed', 'numeric-value-matched', 'declaration-present-abi-unassessed'): raise ValueError(f'unsupported raw assessment: {ident}')
        if row['raw_state'] == 'declaration-present-abi-unassessed' and (not row['exact_name_candidates'] or 'value' in ids[ident]): raise ValueError(f'invalid declaration assessment: {ident}')
        if row['raw_state'] == 'numeric-value-matched':
            symbol = ids[ident]['symbol']
            if symbol not in constants or 'value' not in ids[ident] or constants[symbol] != ids[ident]['value']:
                raise ValueError(f'numeric value mismatch: {ident}')
            if not any(c.get('path') == 'crates/excel-api-sys/src/constants.rs' and c.get('symbol') == symbol for c in row['exact_name_candidates']):
                raise ValueError(f'missing constant evidence: {ident}')
        if not isinstance(row['exact_name_candidates'],list): raise ValueError(f'invalid candidate list: {ident}')
        if row['exact_name_candidates'] != definitions.get(ids[ident]['symbol'], []): raise ValueError(f'missing or extra exact-name candidates: {ident}')
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
        audited = [CATALOGUE, 'crates/excel-api-sys/src', 'crates/excel-api/src', 'crates/excel-api-macros/src']
        stale = subprocess.run(['git', 'diff', '--quiet', REVISION, '--', *audited], cwd=ROOT)
        if stale.returncode != 0:
            raise ValueError('audited source or catalogue differs from pinned revision; reassessment required')
        new = subprocess.check_output(['git', 'ls-files', '--others', '--exclude-standard', '--', *audited], cwd=ROOT)
        if new.strip():
            raise ValueError('untracked files in audited source paths; reassessment required')
        rows = [json.loads(line) for line in DRAFT.read_text(encoding='utf-8').splitlines()]
        print(f'PASS: {verify(rows)} IDs accounted for; candidates and numeric claims verified at {REVISION}')
        print('LIMIT: numeric matches do not prove functional, ABI, test or host coverage')
    except (ValueError, OSError, subprocess.CalledProcessError) as exc:
        print(f'FAIL: {exc}', file=sys.stderr); return 1
    return 0

if __name__ == '__main__': sys.exit(main())
