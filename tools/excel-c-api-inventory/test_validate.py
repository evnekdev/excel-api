import copy
import json
import unittest
import validate as v

class ValidatorTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.seed = v.read_jsonl(v.CATALOGUE.read_bytes())
        cls.baseline = {r['id']: v.identity(r) for r in cls.seed}

    def changed(self, change, expected):
        records = copy.deepcopy(self.seed)
        change(records)
        with self.assertRaisesRegex(v.Invalid, expected):
            v.validate(records, self.baseline)

    def test_seed(self):
        self.assertEqual(v.validate(self.seed, self.baseline), 15)

    def test_json_and_utf8(self):
        for bad in (b'{', b'\xff', b'{"a":1,"a":2}', b'[]'):
            with self.subTest(bad=bad), self.assertRaises(v.Invalid): v.read_jsonl(bad)

    def test_required_and_type(self):
        self.changed(lambda r: r[0].pop('description'), 'description')
        self.changed(lambda r: r[0].update(edna_ids='EDNA-02'), 'edna_ids')

    def test_duplicate_id(self):
        self.changed(lambda r: r[1].update(id=r[0]['id']), 'duplicate ID')

    def test_duplicate_identity(self):
        def mutate(r):
            r[1]['symbol'] = r[0]['symbol']; r[1]['kind'] = r[0]['kind']
        self.changed(mutate, 'duplicate canonical identity')

    def test_reassigned_id(self):
        self.changed(lambda r: r[0].update(symbol='Different'), 'historical ID reassigned')

    def test_edna(self):
        self.changed(lambda r: r[0].update(edna_ids=['EDNA-99']), 'EDNA')
        self.changed(lambda r: r[0].update(edna_ids=['EDNA-02', 'EDNA-02']), 'EDNA')

    def test_source(self):
        r = copy.deepcopy(self.seed)
        r[0]['source']['path'] = 'bad'
        with self.assertRaisesRegex(v.Invalid, 'source'):
            v.validate(r)
        self.changed(lambda r: r[0]['source'].update(blob_sha='0'*40), 'blob')
        self.changed(lambda r: r[0]['source'].update(line_end=99999), 'range')

    def test_comment_is_not_declaration(self):
        r = copy.deepcopy(self.seed[4]); r['source']['line_start'] = 1; r['source']['line_end'] = 1
        with self.assertRaises(v.Invalid): v.declaration(r, '/* #define xltypeNum 1 */')

    def test_macro_value_is_explicitly_unverified(self):
        r = copy.deepcopy(self.seed[4]); r['source']['line_start'] = 1; r['source']['line_end'] = 1
        self.assertTrue(v.declaration(r, '#define xltypeNum 123456'))
        # Schema v1 has no expected value; the CLI must report this limitation.
        from contextlib import redirect_stdout
        from io import StringIO
        out = StringIO()
        with redirect_stdout(out): self.assertEqual(v.main(['check']), 0)
        self.assertIn('macro values/aliases/guards', out.getvalue())

    def test_host_pass_needs_evidence(self):
        self.changed(lambda r: r[0].update(excel_host_status='pass'), 'excel_host_evidence')
        self.changed(lambda r: r[0].update(excel_host_status='pass', excel_host_evidence=['unverified link']), 'cannot be verified')

    def test_reconcile(self):
        r = self.seed[0]; key = {'path': r['source']['path'], 'kind': r['kind'], 'symbol': r['symbol']}
        entry = dict(key, disposition='eligible')
        messages, bad = v.reconcile([], {'complete': True, 'entries': [entry]})
        self.assertTrue(bad); self.assertIn('MISSING', str(messages))
        messages, bad = v.reconcile([r], {'complete': True, 'entries': []})
        self.assertTrue(bad); self.assertIn('EXTRA', str(messages))
        messages, bad = v.reconcile([], {'complete': True, 'entries': [dict(entry, disposition='excluded', reason='disabled')]})
        self.assertFalse(bad); self.assertIn('EXCLUDED', str(messages))
        messages, bad = v.reconcile([], None)
        self.assertFalse(bad); self.assertIn('NOT ASSESSED', str(messages))

    def test_census_expression_and_partial_scope(self):
        r = self.seed[4]
        entry = {'path': r['source']['path'], 'kind': r['kind'],
                 'symbol': r['symbol'], 'disposition': 'eligible',
                 'expression': '0x9999'}
        messages, bad = v.reconcile([r], {'complete': True, 'entries': [entry]})
        self.assertTrue(bad); self.assertIn('CHANGED MACRO EXPRESSION', str(messages))
        messages, bad = v.reconcile([r], {'complete': False, 'entries': []})
        self.assertFalse(bad); self.assertNotIn('EXTRA', str(messages))

    def test_unknown_alias_representation_rejected(self):
        self.changed(lambda r: r[0].update(aliases=['other']), 'unsupported fields')

    def test_freshness_unassessed(self):
        result = '\n'.join(v.freshness(self.seed))
        self.assertIn('Rust implementation: NOT ASSESSED', result)
        self.assertIn('Native Excel evidence: NOT ASSESSED', result)
        self.assertNotIn('needs review', result.lower())
        assessed = copy.deepcopy(self.seed)
        assessed[0]['rust_status'] = 'implemented'
        assessed[0]['rust_evidence'] = ['example evidence']
        from unittest.mock import patch
        with patch.object(v, 'git', side_effect=AssertionError('HEAD must not be queried')):
            self.assertIn('UNKNOWN / insufficient', '\n'.join(v.freshness(assessed)))

if __name__ == '__main__': unittest.main()
