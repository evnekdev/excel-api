import copy
import json
import subprocess
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from io import StringIO
from pathlib import Path
from unittest.mock import patch
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
        messages, bad = v.compare_census([], {'complete': True, 'entries': [entry]})
        self.assertTrue(bad); self.assertIn('MISSING', str(messages))
        messages, bad = v.compare_census([r], {'complete': True, 'entries': []})
        self.assertTrue(bad); self.assertIn('EXTRA', str(messages))
        messages, bad = v.compare_census([], {'complete': True, 'entries': [dict(entry, disposition='excluded', reason='disabled')]})
        self.assertFalse(bad); self.assertIn('EXCLUDED', str(messages))
        messages, bad = v.compare_census([], None)
        self.assertFalse(bad); self.assertIn('NOT ASSESSED', str(messages))

    def test_census_expression_and_partial_scope(self):
        r = self.seed[4]
        entry = {'path': r['source']['path'], 'kind': r['kind'],
                 'symbol': r['symbol'], 'disposition': 'eligible',
                 'expression': '0x9999'}
        messages, bad = v.compare_census([r], {'complete': True, 'entries': [entry]})
        self.assertTrue(bad); self.assertIn('CHANGED MACRO EXPRESSION', str(messages))
        messages, bad = v.compare_census([r], {'complete': False, 'entries': []})
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

    def run_cli(self, records, census='absent', command='reconcile', history=None):
        with tempfile.TemporaryDirectory() as d:
            path = Path(d)
            catalogue = path / 'catalogue.jsonl'
            catalogue.write_text(''.join(json.dumps(r) + '\n' for r in records), encoding='utf8')
            args = [command, '--catalogue', str(catalogue)]
            if census != 'absent':
                file = path / 'census.json'
                file.write_text(json.dumps(census), encoding='utf8')
                args += ['--census', str(file)]
            out, err = StringIO(), StringIO()
            with patch.object(v, 'historical_identities', return_value=history or self.baseline), redirect_stdout(out), redirect_stderr(err):
                code = v.main(args)
            return code, out.getvalue() + err.getvalue()

    def census(self, records, complete=True):
        return {'complete': complete, 'entries': [
            {'path': r['source']['path'], 'kind': r['kind'], 'symbol': r['symbol'], 'disposition': 'eligible'}
            for r in records]}

    def test_reconcile_rejects_invalid_catalogues_before_comparing(self):
        cases = []
        r = copy.deepcopy(self.seed); r[1]['id'] = r[0]['id']; cases.append((r, 'duplicate ID'))
        r = copy.deepcopy(self.seed); r[1] = copy.deepcopy(r[0]); r[1]['id'] = 'CAPI-999999'; cases.append((r, 'duplicate canonical identity'))
        r = copy.deepcopy(self.seed); r[0]['source']['line_end'] = 99999; cases.append((r, 'range'))
        r = copy.deepcopy(self.seed); r[0]['source']['blob_sha'] = '0' * 40; cases.append((r, 'blob'))
        r = copy.deepcopy(self.seed); r[0]['edna_ids'] = ['EDNA-99']; cases.append((r, 'EDNA'))
        r = copy.deepcopy(self.seed); r[0]['id'], r[1]['id'] = r[1]['id'], r[0]['id']; cases.append((r, 'reassigned'))
        for records, reason in cases:
            with self.subTest(reason=reason):
                code, output = self.run_cli(records, self.census(records))
                self.assertEqual(code, 1, output)
                self.assertIn(reason, output)
                self.assertNotIn('MATCHES SUPPLIED CENSUS', output)

    def test_freshness_rejects_invalid_input(self):
        records = copy.deepcopy(self.seed); records[0]['edna_ids'] = ['EDNA-99']
        code, output = self.run_cli(records, command='freshness')
        self.assertEqual(code, 1); self.assertIn('EDNA', output)

    def test_reconcile_exit_statuses_and_findings(self):
        full = self.census(self.seed)
        cases = [('absent', 3, 'NOT ASSESSED'),
                 (self.census(self.seed, False), 3, 'NOT ASSESSED'),
                 (full, 3, 'MATCHES SUPPLIED CENSUS')]
        c = copy.deepcopy(full); c['entries'].append(dict(c['entries'][0], symbol='missing')); cases.append((c, 2, 'MISSING'))
        c = copy.deepcopy(full); c['entries'].pop(); cases.append((c, 2, 'EXTRA'))
        c = copy.deepcopy(full); c['entries'].append(dict(c['entries'][0], symbol='excluded', disposition='excluded', reason='disabled')); cases.append((c, 3, 'EXCLUDED'))
        c = copy.deepcopy(full); c['entries'].append(dict(c['entries'][0], symbol='ambiguous', disposition='ambiguous', reason='guard')); cases.append((c, 2, 'AMBIGUOUS'))
        c = copy.deepcopy(full); c['entries'].append(c['entries'][0]); cases.append((c, 2, 'DUPLICATE'))
        c = copy.deepcopy(full); c['entries'][4]['expression'] = '0x9999'; cases.append((c, 2, 'CHANGED MACRO EXPRESSION'))
        c = copy.deepcopy(full); c['entries'][0]['source_blob'] = '0' * 40; cases.append((c, 2, 'CHANGED SOURCE BLOB'))
        for census, expected, message in cases:
            with self.subTest(message=message):
                code, output = self.run_cli(self.seed, census)
                self.assertEqual(code, expected, output); self.assertIn(message, output)
                if census != 'absent' and expected == 3 and census['complete']:
                    self.assertIn('SOURCE COMPLETENESS NOT VERIFIED', output)

    def test_missing_history_is_not_a_pass(self):
        out, err = StringIO(), StringIO()
        with patch.object(v, 'historical_identities', side_effect=v.HistoryUnavailable('missing baseline')), redirect_stdout(out), redirect_stderr(err):
            code = v.main(['check'])
        self.assertEqual(code, 3); self.assertIn('historical ID protection incomplete', err.getvalue())
        self.assertNotIn('PASS', out.getvalue())

class HistoricalGitTests(unittest.TestCase):
    def test_later_ids_removal_reassignment_and_reuse(self):
        seed = v.read_jsonl(v.CATALOGUE.read_bytes())
        additional = copy.deepcopy(seed[7])
        additional['id'] = 'CAPI-000100'; additional['symbol'] = 'xlStack'
        additional['source']['line_start'] = additional['source']['line_end'] = 382
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            def run(*args):
                return subprocess.check_output(['git', '-C', d, *args], stderr=subprocess.STDOUT).decode().strip()
            run('init', '-q'); run('config', 'user.name', 'Synthetic test'); run('config', 'user.email', 'synthetic@example.invalid')
            path = root / v.CATALOGUE_PATH; path.parent.mkdir(parents=True)
            def commit(rows, message):
                path.write_text(''.join(json.dumps(r) + '\n' for r in rows), encoding='utf8')
                run('add', v.CATALOGUE_PATH); run('commit', '-q', '-m', message)
                return run('rev-parse', 'HEAD')
            first = commit(seed, 'original seed')
            second = commit(seed + [additional], 'add later identity')
            reassigned = copy.deepcopy(additional); reassigned['symbol'] = 'xlSheetId'
            reassigned['source']['line_start'] = reassigned['source']['line_end'] = 385
            third = commit(seed + [reassigned], 'reassign later identity')
            run('switch', '-q', '-c', 'remove-reuse', second)
            fourth = commit(seed, 'remove later identity')
            fifth = commit(seed + [reassigned], 'reuse removed identity')
            with patch.object(v, 'ROOT', root):
                original = v.historical_identities(baseline=first, head=first, upstream=first)
                self.assertEqual(len(original), 15)
                later = v.historical_identities(baseline=first, head=second, upstream=second)
                self.assertEqual(later['CAPI-000100'], v.identity(additional))
                upstream_later = v.historical_identities(baseline=first, head=first, upstream=second)
                self.assertIn('CAPI-000100', upstream_later)
                with self.assertRaisesRegex(v.Invalid, 'reassigned'):
                    v.historical_identities(baseline=first, head=third, upstream=third)
                with self.assertRaises(v.Invalid):
                    v.historical_identities(baseline=first, head=fourth, upstream=fourth)
                with self.assertRaises(v.Invalid):
                    v.historical_identities(baseline=first, head=fifth, upstream=fifth)
                with self.assertRaises(v.HistoryUnavailable):
                    v.historical_identities(baseline='0' * 40, head=second, upstream=second)
            with self.assertRaisesRegex(v.Invalid, 'removed'):
                v.validate(seed, later)
            self.assertEqual(v.validate(seed + [additional], later), 16)
            with self.assertRaisesRegex(v.Invalid, 'reassigned'):
                v.validate(seed + [reassigned], later)

if __name__ == '__main__': unittest.main()
