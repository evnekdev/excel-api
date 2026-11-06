import copy
import json
import unittest
from unittest.mock import patch
import audit_source
import populate
import source_census as source
import validate

class CensusTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.text = source.pinned_header()
        cls.census = source.census(cls.text)
        cls.candidates = {(x['kind'], x['symbol']): x for x in cls.census['entries']}
        cls.records = validate.read_jsonl(validate.CATALOGUE.read_bytes())
        cls.by_symbol = {r['symbol']: r for r in cls.records}

    def test_pinned_source_and_independent_lexical_accounting(self):
        self.assertEqual(self.census['source']['blob_sha'], source.BLOB)
        self.assertEqual(self.census['source']['sha256'], 'd69f860ca5f7be9779f69423276808845ebb2f9648e38aadc99908277a710c74')
        result = audit_source.audit(self.text, self.census)
        self.assertEqual(result['physical_lines'], 1480)
        self.assertEqual(result['raw_macro_names'], 1094)
        self.assertEqual(result['typedef_statements'], 16)
        self.assertEqual(result['prototype_statements'], 6)
        self.assertEqual(result['differences'], [])

    def test_all_candidate_categories_validate(self):
        self.assertEqual(validate.validate(self.records), len(self.candidates))
        self.assertEqual(len(self.records), 1132)
        self.assertEqual(len({validate.identity(r) for r in self.records}), len(self.records))
        self.assertEqual({(r['kind'], r['symbol']) for r in self.records}, set(self.candidates))

    def test_original_identity_and_source_anchor_preserved(self):
        old = validate.read_jsonl(validate.git('show', f'{validate.BASELINE}:{validate.CATALOGUE_PATH}'))
        current = {r['id']: r for r in self.records}
        for r in old:
            self.assertEqual(validate.identity(current[r['id']]), validate.identity(r))
            self.assertEqual(current[r['id']]['source'], r['source'])
            for field, value in r.items():
                self.assertEqual(current[r['id']][field], value, (r['id'], field))

    def test_macro_values_and_classification(self):
        for symbol, kind, value in [('xlflowHalt','flow-code',1), ('xlCommand','function-modifier',0x8000),
                                    ('xlFree','special-function-id',0x4000), ('xlfCount','worksheet-function-id',0),
                                    ('xlcOpen','macro-command-id',0x8001), ('xltypeBigData','type-flag',0x0802),
                                    ('xlHpcRetCallFailed','other-constant',-2)]:
            r = self.by_symbol[symbol]
            self.assertEqual((r['kind'],r['value']), (kind,value))
        bad = copy.deepcopy(self.by_symbol['xltypeBigData']); bad['value'] += 1
        with self.assertRaisesRegex(validate.Invalid, 'effective value mismatch'): validate.validate([bad])
        bad = copy.deepcopy(self.by_symbol['xlFree']); bad['expression'] = '(0 | xlCommand)'
        with self.assertRaisesRegex(validate.Invalid, 'detail or effective value mismatch'): validate.validate([bad])

    def test_safe_expression_limits_and_cycles(self):
        defs = {'A':'1', 'B':'(A | 0x10)', 'C':'B', 'D':'D'}
        self.assertEqual(source.value_of('B', defs), 17)
        self.assertEqual(source.value_of('-2', defs), -2)
        for expr in ['MISSING', 'D', '1 << 8', '__import__("os")']:
            with self.subTest(expr=expr), self.assertRaises(source.ExtractionError): source.value_of(expr, defs)

    def test_unhandled_source_constructs_fail_closed(self):
        cases = [
            self.text.replace('#define xlfCount 0', '#define xlfCount(x) (x)', 1),
            self.text.replace('#define xlfCount 0', '#define xlfCount \\\n  0', 1),
            self.text + '\n#define xlfCount 0\n',
            self.text.replace('#define xlfCount 0', '#define xlfCount MISSING', 1),
            self.text + '\n#ifdef BAD\n#else\n#else\n#endif\n',
        ]
        for changed in cases:
            with self.subTest(tail=changed[-30:]), self.assertRaises(source.ExtractionError):
                source.extract(changed)

    def test_explicit_aliases_are_not_numeric_collisions(self):
        defs = {'A':'1', 'B':'1', 'C':'(A)'}
        self.assertEqual(source.explicit_alias(defs['C'], defs), 'A')
        self.assertIsNone(source.explicit_alias(defs['B'], defs))
        self.assertEqual(self.by_symbol['xltypeNum']['value'], self.by_symbol['xlflowHalt']['value'])
        self.assertNotIn('alias_of', self.by_symbol['xlflowHalt'])
        self.assertFalse([c for c in self.candidates.values() if c.get('alias_of')])

    def test_typedefs_pointers_tags_and_callback(self):
        for kind, name in [('scalar-typedef','BOOL'),('struct','XLREF'),('struct-tag','xlref'),
                           ('pointer-typedef','LPXLREF'),('function-pointer-typedef','PXL_HPC_ASYNC_CALLBACK')]:
            self.assertIn((kind,name), self.candidates)
        self.assertEqual(self.by_symbol['BOOL']['underlying'], 'INT32')
        self.assertEqual(self.by_symbol['LPXLREF']['target'], 'XLREF')
        self.assertIn('CALLBACK', self.by_symbol['PXL_HPC_ASYNC_CALLBACK']['signature'])
        for name in ('Excel4','Excel4v','Excel12','Excel12v','XLCallVer','LPenHelper'):
            self.assertIn('extern "C"', self.by_symbol[name]['linkage'])
        self.assertIn('...', self.by_symbol['Excel4']['signature'])
        self.assertIn('LPXLOPER12', self.by_symbol['Excel12v']['signature'])

    def test_conditional_members_and_guards(self):
        fields = self.by_symbol['XLOPER']['conditional_fields']
        self.assertEqual(fields, [{'guard':'__cplusplus','declaration':'WORD xbool'}, {'guard':'!__cplusplus','declaration':'WORD bool'}])
        changed = self.text.replace('#define xlfCount 0', '#ifdef SAMPLE\n#define xlfCount 0\n#endif', 1)
        candidate = next(c for c in source.extract(changed) if c['symbol'] == 'xlfCount')
        self.assertEqual(candidate['guard'], ['SAMPLE'])
        bad = copy.deepcopy(self.by_symbol['XLOPER']); bad['conditional_fields'][0]['declaration'] = 'WORD bool'
        with self.assertRaises(validate.Invalid): validate.validate([bad])
        with self.assertRaises(source.ExtractionError): source.extract(self.text + '\n#if 0\n#define ABC 1\n#endif\n')

    def test_comment_only_and_fault_injection(self):
        observed = copy.deepcopy(self.census)
        observed['entries'] = [x for x in observed['entries'] if x['symbol'] != 'xlfCount']
        self.assertTrue(audit_source.audit(self.text, observed)['differences'])
        extra = self.text + '\n/* #define NOT_A_REAL_SYMBOL 1 */\n'
        self.assertNotIn('NOT_A_REAL_SYMBOL', [x['symbol'] for x in source.extract(extra)])
        changed = self.text.replace('#define xlfCount 0', '#define xlfCount 1', 1)
        self.assertIn('effective value mismatch', str(audit_source.audit(changed, self.census)['differences']))
        row = self.by_symbol['xlfCount']
        with self.assertRaisesRegex(validate.Invalid, 'effective value mismatch'):
            validate.validate([row], blob_reader=lambda s: changed)

    def test_census_mismatches_and_no_completeness_claim(self):
        c = copy.deepcopy(self.census); c['entries'] = [x for x in c['entries'] if x['symbol'] != 'xlfCount']
        messages, bad = validate.compare_census(self.records, c)
        self.assertTrue(bad); self.assertIn('EXTRA', str(messages))
        c = copy.deepcopy(self.census); c['entries'][0]['symbol'] = 'unlisted'
        messages, bad = validate.compare_census(self.records, c)
        self.assertTrue(bad); self.assertIn('MISSING', str(messages))
        c = copy.deepcopy(self.census); item = next(x for x in c['entries'] if x['symbol'] == 'xlfCount'); item['value'] = 99
        messages, bad = validate.compare_census(self.records, c)
        self.assertTrue(bad); self.assertIn('CHANGED MACRO VALUE', str(messages))
        del item['expression']
        messages, bad = validate.compare_census(self.records, c)
        self.assertTrue(bad); self.assertIn('CHANGED MACRO VALUE', str(messages))
        c = copy.deepcopy(self.census); item = next(x for x in c['entries'] if x['symbol'] == 'LPenHelper'); item['signature'] = 'void LPenHelper(void)'
        messages, bad = validate.compare_census(self.records, c)
        self.assertTrue(bad); self.assertIn('CHANGED SOURCE signature', str(messages))

    def test_generation_is_idempotent_and_history_protected(self):
        history = validate.historical_identities()
        result = populate.build(copy.deepcopy(self.records), source.extract(self.text), history)
        self.assertEqual(result, self.records)
        validate.validate(result, history)

if __name__ == '__main__': unittest.main()
