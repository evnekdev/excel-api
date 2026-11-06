"""Validate the preliminary #122 candidate inventory, not implementation coverage."""
import ast
import collections
import functools
import json
import pathlib
import re
import subprocess
import sys

import validate as catalogue_validator

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
    'xleventCalculationCanceled': ('crates/excel-api/src/runtime.rs', 'async_events_register_once_across_close_and_reopen'),
    'xleventCalculationEnded': ('crates/excel-api/src/runtime.rs', 'async_events_register_once_across_close_and_reopen'),
    'xlfNow': ('crates/excel-api-sys/src/constants.rs', 'on_time_and_now_match_checked_in_xlcall_header'),
    'xlcOnTime': ('crates/excel-api-sys/src/constants.rs', 'on_time_and_now_match_checked_in_xlcall_header'),
    'xlretAbort': ('crates/excel-api/src/excel_call.rs', 'abort_is_available_to_verified_contexts_and_preserves_return_codes'),
    'xlretUncalced': ('crates/excel-api/src/excel_call.rs', 'abort_is_available_to_verified_contexts_and_preserves_return_codes'),
    'xlretInvAsynchronousContext': ('crates/excel-api/src/async_udf.rs', 'excel_return_code_and_boolean_rejection_are_preserved'),
}
ERROR_VARIANTS = dict(zip(
    ('xlerrNull','xlerrDiv0','xlerrValue','xlerrRef','xlerrName','xlerrNum','xlerrNA','xlerrGettingData'),
    ('Null','Div0','Value','Ref','Name','Num','Na','GettingData')))
BORROWED_WRAPPERS = {'XLOPER12': ('RawExcelValue', 'raw: NonNull<XLOPER12>'),
                     'XLREF12': ('ExcelReferenceArea', 'area: NonNull<XLREF12>'),
                     'XLMREF12': ('ExcelMultiReference', 'areas: NonNull<XLREF12>'),
                     'XCHAR': ('ExcelStr', 'payload: NonNull<XCHAR>'),
                     'IDSHEET': ('ExcelMultiReference', 'sheet_id: IDSHEET')}
ABI_CHECKED = {'BOOL','XCHAR','RW','COL','IDSHEET','XLREF12','XLMREF12','FP12','XLOPER12'}
LEGACY_NAMES = {'XLREF','xlref','LPXLREF','XLMREF','xlmref','LPXLMREF',
                'FP','_FP','XLOPER','xloper','LPXLOPER','Excel4','Excel4v'}

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

def core_symbol_occurrences():
    paths = subprocess.check_output(
        ['git', 'ls-tree', '-r', '--name-only', REVISION, '--',
         'crates/excel-api-sys/src', 'crates/excel-api/src', 'crates/excel-api-macros/src'],
        cwd=ROOT, text=True, encoding='utf8').splitlines()
    counts = collections.Counter()
    for path in paths:
        if path.endswith('.rs'):
            counts.update(re.findall(r'[A-Za-z_][A-Za-z_0-9]*', git_show(path)))
    return counts

def supplemental_mentions(names):
    """Bounded lexical leads in tests, examples and the ABI checker, not semantic evidence."""
    scopes = ['crates/excel-api-sys/tests', 'crates/excel-api/tests',
              'crates/excel-api-macros/tests', 'examples/minimal-xll/src',
              'examples/minimal-xll/tests', 'tools/abi-check/src', 'tools/abi-check/native']
    paths = subprocess.check_output(['git', 'ls-tree', '-r', '--name-only', REVISION, '--', *scopes], cwd=ROOT, text=True, encoding='utf8').splitlines()
    found = collections.defaultdict(list)
    for path in paths:
        if not path.endswith(('.rs', '.cpp', '.h', '.snap')):
            continue
        for number, line in enumerate(git_show(path).splitlines(), 1):
            for token in set(re.findall(r'[A-Za-z_][A-Za-z_0-9]*', line)) & names:
                found[token].append({'path': path, 'line': number})
    return found

def abi_numeric_checks():
    rust = git_show('tools/abi-check/src/main.rs')
    native = git_show('tools/abi-check/native/layout_check.cpp')
    checks = {}
    pattern = r'check!\(\s*"(\w+)"\s*,\s*(\w+)\s*,\s*(excel_sdk_\w+)\(\)\s*\)'
    for match in re.finditer(pattern, rust):
        label, symbol, helper = match.groups()
        if label == symbol and re.search(r'CONST_[UI]32_FN\(\s*' + re.escape(helper) + r'\s*,\s*' + re.escape(symbol) + r'\s*\)', native):
            checks[symbol] = rust.count('\n', 0, match.start()) + 1
    return checks

def call_descriptor_names():
    source = git_show('crates/excel-api/src/excel_call.rs')
    pattern = r'pub const \w+: ExcelCallDescriptor = ExcelCallDescriptor \{(.*?)\n\};'
    result = set()
    for block in re.findall(pattern, source, flags=re.DOTALL):
        match = re.search(r'\bname: "(\w+)"', block)
        if not match:
            raise ValueError('descriptor without a literal name; update audit')
        result.add(match.group(1))
    return result

def edna_relations(item):
    """Preserve canonical links and label bounded source/code-derived candidates."""
    candidates = set()
    kind, symbol = item['kind'], item['symbol']
    if kind in ('type-flag', 'error-code'):
        candidates.add('EDNA-03')
    if kind == 'ownership-flag':
        candidates.add('EDNA-04')
    if kind == 'return-code':
        candidates.add('EDNA-06')
    if kind == 'event-id':
        candidates.add('EDNA-14')
    if symbol in ('xltypeRef','xltypeSRef','XLREF12','XLMREF12','LPXLREF12','LPXLMREF12','xlref12','xlmref12'):
        candidates.add('EDNA-05')
    if symbol in ('XLOPER12','FP12','xloper12','_FP12'):
        candidates.add('EDNA-03')
    if symbol in ('Excel12','Excel12v'):
        candidates.add('EDNA-06')
    if symbol in ('xlfRegister','xlfUnregister','xlfRegisterId'):
        candidates.add('EDNA-07')
    if symbol == 'xlFree': candidates.add('EDNA-04')
    if symbol == 'xlAsyncReturn': candidates.add('EDNA-11')
    if symbol == 'xlEventRegister': candidates.add('EDNA-14')
    canonical = set(item['edna_ids'])
    return ([{'id': ident, 'basis': 'canonical-source-link'} for ident in sorted(canonical)] +
            [{'id': ident, 'basis': 'audit-candidate'} for ident in sorted(candidates - canonical)])

def verify(rows):
    catalogue = [json.loads(line) for line in git_show(CATALOGUE).splitlines()]
    ids = {row['id']: row for row in catalogue}
    if len(ids) != len(catalogue): raise ValueError('duplicate source catalogue ID')
    found = set()
    aggregate_ids = {item['symbol']: item['id'] for item in catalogue if item['kind'] == 'struct'}
    constants = numeric_constants()
    abi_numbers = abi_numeric_checks()
    descriptors = call_descriptor_names()
    known_edna = set(re.findall(r'\*\*(EDNA-\d{2})\*\*', git_show('docs/architecture/excel-dna-subsystems-and-crates.md')))
    definitions = public_definitions()
    occurrences = core_symbol_occurrences()
    mentions = supplemental_mentions({item['symbol'] for item in catalogue})
    allowed = {'crates/excel-api-sys/src/constants.rs', 'crates/excel-api-sys/src/types.rs', 'crates/excel-api-sys/src/functions.rs'}
    cache = {}
    for row in rows:
        if not isinstance(row, dict): raise ValueError('assessment row must be an object')
        if set(row) != {'assessment_version','rust_revision','capi_id','raw_state','functional_state','test_state','host_state','exact_name_candidates','functional_evidence','related_rust','test_evidence','abi_evidence','descriptor_observation','related_evidence','core_name_occurrences','supplemental_mentions','dispatch_relation','type_target_capi_id','edna_relations'}: raise ValueError('invalid row fields')
        if type(row['assessment_version']) != int or row['assessment_version'] != 1: raise ValueError('unsupported assessment version')
        if row['rust_revision'] != REVISION: raise ValueError('incorrect audited Rust revision')
        ident = row['capi_id']
        if not isinstance(ident, str): raise ValueError('CAPI ID must be a string')
        if ident not in ids or ident in found: raise ValueError(f'duplicate or invalid ID: {ident}')
        found.add(ident)
        item = ids[ident]
        if row['edna_relations'] != edna_relations(item) or any(relation['id'] not in known_edna for relation in row['edna_relations']): raise ValueError(f'invalid EDNA relation: {ident}')
        expected_target = aggregate_ids.get(item.get('target')) if item['kind'] in ('pointer-typedef', 'struct-tag') else None
        if row['type_target_capi_id'] != expected_target or (item['kind'] in ('pointer-typedef', 'struct-tag') and expected_target is None): raise ValueError(f'invalid type relationship: {ident}')
        if type(row['core_name_occurrences']) != int or row['core_name_occurrences'] != occurrences[ids[ident]['symbol']]: raise ValueError(f'invalid core occurrence count: {ident}')
        hits = mentions.get(ids[ident]['symbol'], [])
        if row['supplemental_mentions'] != {'total_lines': len(hits), 'first_three': hits[:3]}: raise ValueError(f'invalid supplemental mentions: {ident}')
        additional = row['related_evidence']
        if not isinstance(additional, list): raise ValueError(f'invalid related evidence: {ident}')
        evidence_keys = set()
        for evidence in additional:
            if not isinstance(evidence, dict) or set(evidence) != {'path','line','role'} or not isinstance(evidence['path'], str) or not isinstance(evidence['role'], str) or type(evidence['line']) != int: raise ValueError(f'invalid related evidence: {ident}')
            key = (evidence['path'], evidence['line'], evidence['role'])
            if key in evidence_keys: raise ValueError(f'duplicate related evidence: {ident}')
            evidence_keys.add(key)
            n = evidence['line']; symbol = ids[ident]['symbol']
            if evidence['role'] == 'return-materialization' and evidence['path'] == 'crates/excel-api/src/return_alloc.rs':
                lines = git_show(evidence['path']).splitlines()
                if not 1 <= n < 800 or not re.search(r'\b' + re.escape(symbol) + r'\b', lines[n-1]): raise ValueError(f'stale return evidence: {ident}')
            elif evidence['role'] == 'owned-conversion-boundary' and evidence['path'] == 'crates/excel-api/src/convert.rs':
                lines = git_show(evidence['path']).splitlines()
                expected = 'ExcelValueRef::Reference(_) => Err(ConversionError::UnsupportedReference),' if symbol in ('xltypeRef','xltypeSRef') else 'return Err(ConversionError::NestedArrayUnsupported);' if symbol == 'xltypeMulti' else None
                if expected is None or not 1 <= n <= len(lines) or lines[n-1].strip() != expected: raise ValueError(f'stale conversion evidence: {ident}')
            else:
                raise ValueError(f'unsupported related evidence: {ident}')
        if not isinstance(row['exact_name_candidates'], list) or not all(isinstance(candidate, dict) for candidate in row['exact_name_candidates']): raise ValueError(f'invalid candidate list: {ident}')
        item = ids[ident]
        dispatch = None
        if item['symbol'] in ('xlGetFmlaInfo','xlGetMouseInfo'):
            dispatch = 'lpenhelper-selector-only'
        elif item['kind'] in ('worksheet-function-id','macro-command-id','special-function-id') or item['symbol'] == 'xlUDF':
            dispatch = 'generic-selector-availability-unassessed'
        if row['dispatch_relation'] != dispatch: raise ValueError(f'invalid source dispatch relation: {ident}')
        applicable = item['kind'] in ('worksheet-function-id', 'macro-command-id', 'special-function-id')
        expected_observation = ('present' if item['symbol'] in descriptors else 'not-observed') if applicable else None
        if row['descriptor_observation'] != expected_observation: raise ValueError(f'wrong descriptor observation: {ident}')
        abi = row['abi_evidence']; symbol = ids[ident]['symbol']
        if abi is not None:
            if not isinstance(abi, dict) or set(abi) != {'path','line','target','command','check'} or abi['path'] not in ('tools/abi-check/src/main.rs','tools/abi-check/native/layout_check.cpp') or abi['target'] != 'x86_64-pc-windows-msvc' or abi['command'] != 'cargo run --manifest-path tools/abi-check/Cargo.toml' or type(abi['line']) != int: raise ValueError(f'invalid ABI evidence: {ident}')
            lines = git_show(abi['path']).splitlines(); n = abi['line']
            if abi['path'] == 'tools/abi-check/src/main.rs' and abi['check'] == 'size' and symbol in ABI_CHECKED:
                if not 1 <= n <= len(lines) or f'size_of::<{symbol}>()' not in lines[n-1]: raise ValueError(f'stale ABI size evidence: {ident}')
            elif abi['path'] == 'tools/abi-check/src/main.rs' and abi['check'] == 'numeric-value' and abi_numbers.get(symbol) == n and row['raw_state'] == 'numeric-value-matched':
                pass
            elif abi['path'] == 'tools/abi-check/native/layout_check.cpp' and abi['check'] == 'c-prototype-only' and symbol in ('Excel12','Excel12v') and 1 <= n <= len(lines) and f'decltype(&{symbol})' in lines[n-1]:
                pass
            else:
                raise ValueError(f'stale ABI evidence: {ident}')
        related = row['related_rust']
        if related is not None:
            if not isinstance(related, dict) or set(related) != {'path', 'line', 'symbol'} or not isinstance(related['symbol'], str) or related['symbol'] != ALTERNATE_NAMES.get(ids[ident]['symbol']) or not isinstance(related['path'], str) or related['path'] not in {'crates/excel-api-sys/src/types.rs', 'crates/excel-api-sys/src/functions.rs'} or type(related['line']) != int: raise ValueError(f'invalid related declaration: {ident}')
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
        elif row['test_state'] == 'abi-check-passed-this-run':
            if test is not None or abi is None: raise ValueError(f'missing ABI test evidence: {ident}')
        elif row['test_state'] == 'compile-check-passed-this-run':
            if ids[ident]['symbol'] != 'LPXLOPER12' or not isinstance(test, dict) or set(test) != {'path','line','symbol','configuration','run'} or test['path'] != 'crates/excel-api/tests/macro_metadata.rs' or test['symbol'] != 'LPXLOPER12' or test['configuration'] != 'default' or test['run'] != 'cargo test -p excel-api --test macro_metadata' or type(test['line']) != int: raise ValueError(f'invalid compile evidence: {ident}')
            lines = git_show(test['path']).splitlines(); n = test['line']
            if not 1 <= n <= len(lines) or not lines[n-1].startswith('const _: unsafe extern "system" fn(') or 'LPXLOPER12' not in lines[n-1]: raise ValueError(f'stale compile evidence: {ident}')
        elif row['test_state'] != 'unassessed' or test is not None: raise ValueError(f'unsupported test assessment: {ident}')
        if row['functional_state'] not in ('unassessed', 'descriptor-present', 'decode-branch-present', 'decoding-explicitly-unsupported', 'async-return-operation-present', 'ownership-mask-present', 'error-conversion-present', 'event-registration-present', 'borrowed-wrapper-present', 'fp12-registration-type-not-observed', 'generic-return-code-wrapper', 'reference-coordinate-accessor-present', 'boolean-decoding-present'): raise ValueError(f'unsupported functional assessment: {ident}')
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
        elif row['functional_state'] == 'boolean-decoding-present':
            evidence = row['functional_evidence']
            if ids[ident]['symbol'] != 'BOOL' or not isinstance(evidence, dict) or set(evidence) != {'path','line','symbol','configuration'} or evidence['path'] != 'crates/excel-api/src/borrowed.rs' or evidence['symbol'] != 'XLTYPE_BOOL' or evidence['configuration'] != 'default' or type(evidence['line']) != int: raise ValueError(f'invalid Boolean evidence: {ident}')
            lines = git_show(evidence['path']).splitlines(); n = evidence['line']
            if not 1 <= n <= len(lines) or lines[n-1].strip() != 'XLTYPE_BOOL => {' or 'ExcelValueRef::Boolean(value != 0)' not in '\n'.join(lines[n:n+6]) or 'pub xbool: BOOL,' not in git_show('crates/excel-api-sys/src/types.rs'): raise ValueError(f'stale Boolean evidence: {ident}')
        elif row['functional_state'] == 'reference-coordinate-accessor-present':
            evidence = row['functional_evidence']; symbol = ids[ident]['symbol']
            accessor = {'RW': ('first_row', 'rwFirst'), 'COL': ('first_column', 'colFirst')}.get(symbol)
            if not accessor or not isinstance(evidence, dict) or set(evidence) != {'path','line','symbol','configuration'} or evidence['path'] != 'crates/excel-api/src/borrowed.rs' or evidence['symbol'] != accessor[0] or evidence['configuration'] != 'default' or type(evidence['line']) != int: raise ValueError(f'invalid coordinate evidence: {ident}')
            lines = git_show(evidence['path']).splitlines(); n = evidence['line']
            if not 1 <= n < len(lines) or lines[n-1].strip() != f'pub fn {accessor[0]}(&self) -> {symbol} {{' or lines[n].strip() != f'self.read().{accessor[1]}': raise ValueError(f'stale coordinate evidence: {ident}')
        elif row['functional_state'] == 'generic-return-code-wrapper':
            evidence = row['functional_evidence']
            if ids[ident]['kind'] != 'return-code' or row['raw_state'] != 'numeric-value-matched' or not isinstance(evidence, dict) or set(evidence) != {'path','line','symbol','configuration'} or evidence['path'] != 'crates/excel-api/src/excel_call.rs' or evidence['symbol'] != 'ExcelReturnCode' or evidence['configuration'] != 'default' or type(evidence['line']) != int: raise ValueError(f'invalid return code evidence: {ident}')
            lines = git_show(evidence['path']).splitlines(); n = evidence['line']
            if not 1 <= n <= len(lines) or lines[n-1] != 'pub struct ExcelReturnCode(pub i32);' or 'self.0 & flag != 0' not in '\n'.join(lines[n:n+15]) or 'let code = ExcelReturnCode(raw);' not in '\n'.join(lines): raise ValueError(f'stale return code evidence: {ident}')
        elif row['functional_state'] == 'fp12-registration-type-not-observed':
            evidence = row['functional_evidence']
            if ids[ident]['symbol'] != 'FP12' or not isinstance(evidence, dict) or set(evidence) != {'path','line','symbol','configuration'} or evidence['path'] != 'crates/excel-api/src/registration.rs' or evidence['symbol'] != 'ExcelArgumentType' or evidence['configuration'] != 'default' or type(evidence['line']) != int: raise ValueError(f'invalid FP12 evidence: {ident}')
            lines = git_show(evidence['path']).splitlines(); n = evidence['line']
            if not 1 <= n <= len(lines) or lines[n-1] != 'pub enum ExcelArgumentType {' or 'pub enum ExcelReturnType {' not in lines: raise ValueError(f'stale FP12 evidence: {ident}')
            for enum in ('ExcelArgumentType', 'ExcelReturnType'):
                begin = lines.index(f'pub enum {enum} {{'); end = lines.index('}', begin)
                if any('FP12' in line.upper() for line in lines[begin:end]): raise ValueError(f'FP12 registration exists: {ident}')
        elif row['functional_state'] == 'borrowed-wrapper-present':
            evidence = row['functional_evidence']; pair = BORROWED_WRAPPERS.get(ids[ident]['symbol'])
            if not pair or not isinstance(evidence, dict) or set(evidence) != {'path','line','symbol','configuration'} or evidence['path'] != 'crates/excel-api/src/borrowed.rs' or evidence['symbol'] != pair[0] or evidence['configuration'] != 'default' or type(evidence['line']) != int: raise ValueError(f'invalid wrapper evidence: {ident}')
            lines = git_show(evidence['path']).splitlines(); n = evidence['line']
            if not 1 <= n <= len(lines) or not lines[n-1].startswith(f'pub struct {pair[0]}<') or pair[1] not in '\n'.join(lines[n-1:n+5]): raise ValueError(f'stale wrapper evidence: {ident}')
            if ids[ident]['symbol'] == 'XLMREF12':
                types_source = git_show('crates/excel-api-sys/src/types.rs')
                if 'pub type LPXLMREF12 = *mut XLMREF12;' not in types_source or 'pub lpmref: LPXLMREF12,' not in types_source or 'NonNull::new(value.lpmref)' not in '\n'.join(lines): raise ValueError(f'missing multi-reference decode path: {ident}')
        elif row['functional_state'] == 'event-registration-present':
            evidence = row['functional_evidence']; symbol = ids[ident]['symbol']
            if symbol not in ('xleventCalculationCanceled','xleventCalculationEnded') or not isinstance(evidence, dict) or set(evidence) != {'path','line','symbol','configuration'} or evidence['path'] != 'crates/excel-api/src/runtime.rs' or evidence['symbol'] != 'ensure_async_events' or evidence['configuration'] != 'async-functions' or type(evidence['line']) != int: raise ValueError(f'invalid event evidence: {ident}')
            lines = git_show(evidence['path']).splitlines(); n = evidence['line']
            if not 1 <= n <= len(lines) or lines[n-1].strip() != f'excel_api_sys::{symbol},' or 'context.register_event(' not in '\n'.join(lines[max(0,n-5):n]): raise ValueError(f'stale event evidence: {ident}')
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
        if row['raw_state'] not in ('unassessed', 'numeric-value-matched', 'declaration-present-abi-unassessed', 'related-declaration-abi-unassessed', 'legacy-public-declaration-not-observed'): raise ValueError(f'unsupported raw assessment: {ident}')
        if row['raw_state'] == 'legacy-public-declaration-not-observed':
            if ids[ident]['symbol'] not in LEGACY_NAMES or row['exact_name_candidates'] or related is not None or 'Legacy XLOPER structures and Excel4 are unsupported.' not in git_show('crates/excel-api-sys/src/constants.rs'): raise ValueError(f'invalid legacy-name assessment: {ident}')
        if row['raw_state'] == 'related-declaration-abi-unassessed' and related is None: raise ValueError(f'missing related declaration: {ident}')
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
            if set(candidate) != {'path','line','symbol'} or not isinstance(candidate['path'], str) or candidate['path'] not in allowed or candidate['symbol'] != ids[ident]['symbol'] or type(candidate['line']) != int: raise ValueError(f'invalid candidate: {ident}')
            key = (candidate['path'], candidate['line'], candidate['symbol'])
            if key in candidate_keys: raise ValueError(f'duplicate candidate: {ident}')
            candidate_keys.add(key)
            path = candidate['path']
            if path not in cache: cache[path] = git_show(path).splitlines()
            line = candidate['line']
            if not 1 <= line <= len(cache[path]) or not re.match(r'\s*pub\s+(?:const|struct|union|enum|type)\s+' + re.escape(candidate['symbol']) + r'\b', cache[path][line-1]): raise ValueError(f'stale candidate: {ident}')
    if found != set(ids): raise ValueError(f'missing IDs: {len(set(ids)-found)}')
    return len(found)

def verify_current_inputs():
    header = 'tools/Excel2013XLLSDK/INCLUDE/XLCALL.H'
    disk_header = (ROOT / header).read_bytes().replace(b'\r\n', b'\n')
    try:
        pinned_header = subprocess.check_output(['git', 'show', f'6c8354d6536e18b58559f9cc3134f25a9e58d478:{header}'], cwd=ROOT, stderr=subprocess.PIPE)
    except subprocess.CalledProcessError as exc:
        raise catalogue_validator.SourceUnavailable('pinned SDK source object unavailable') from exc
    if disk_header != pinned_header:
        raise ValueError('working SDK header differs from pinned blob after CRLF normalization')
    try:
        subprocess.check_output(['git', 'cat-file', '-e', f'{REVISION}^{{commit}}'], cwd=ROOT, stderr=subprocess.PIPE)
    except subprocess.CalledProcessError as exc:
        raise catalogue_validator.HistoryUnavailable('audited Rust revision unavailable') from exc
    audited = [CATALOGUE, 'docs/architecture/excel-dna-subsystems-and-crates.md', 'docs/evidence/excel-c-api-source-census.md', 'crates/excel-api-sys/src', 'crates/excel-api-sys/tests',
               'crates/excel-api/src', 'crates/excel-api/tests',
               'crates/excel-api-macros/src', 'crates/excel-api-macros/tests',
               'examples/minimal-xll/src', 'examples/minimal-xll/tests', 'tools/abi-check']
    stale = subprocess.run(['git', 'diff', '--quiet', REVISION, '--', *audited], cwd=ROOT)
    if stale.returncode != 0:
        raise ValueError('audited source or catalogue differs from pinned revision; reassessment required')
    new = subprocess.check_output(['git', 'ls-files', '--others', '--exclude-standard', '--', *audited], cwd=ROOT)
    if new.strip():
        raise ValueError('untracked files in audited source paths; reassessment required')
    catalogue_validator.validate(catalogue_validator.read_jsonl((ROOT / CATALOGUE).read_bytes()),
                                 catalogue_validator.historical_identities())

def main():
    try:
        verify_current_inputs()
        rows = catalogue_validator.read_jsonl(DRAFT.read_bytes())
        print(f'PASS: {verify(rows)} IDs accounted for; candidates and numeric claims verified at {REVISION}')
        print('LIMIT: numeric matches do not prove functional, ABI, test or host coverage')
    except (catalogue_validator.HistoryUnavailable, catalogue_validator.SourceUnavailable) as exc:
        print(f'NOT VERIFIED: {exc}', file=sys.stderr); return 3
    except (ValueError, OSError, subprocess.CalledProcessError) as exc:
        print(f'FAIL: {exc}', file=sys.stderr); return 1
    return 0

if __name__ == '__main__': sys.exit(main())
