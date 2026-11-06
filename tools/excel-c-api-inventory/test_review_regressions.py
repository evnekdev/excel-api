"""Fault injections found during the independent technical review (#117/#118)."""
import copy
import json
from pathlib import Path
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from io import StringIO
from unittest.mock import patch

import populate
import source_census as source
import validate


class ReviewRegressions(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.text = source.pinned_header()
        cls.rows = validate.read_jsonl(validate.CATALOGUE.read_bytes())

    def test_unsupported_active_source_is_not_silently_omitted(self):
        for addition in ('extern int Unlisted;', 'enum Color { Red };',
                         '#include "another.h"', '#undef xlCommand',
                         'typedef unsigned long Unlisted;', '#pragma pack(1)',
                         '#define', '#define 42'):
            with self.subTest(addition=addition), self.assertRaises(source.ExtractionError):
                source.extract(self.text + '\n' + addition + '\n')
        with self.assertRaises(source.ExtractionError):
            source.extract(self.text.replace('typedef INT32 BOOL;', 'typedef INT32 BOOL; extern int Hidden;', 1))

    def test_typedef_guards_are_preserved(self):
        text = self.text.replace('typedef INT32 BOOL;', '#ifdef EXTRA\ntypedef INT32 BOOL;\n#endif', 1)
        row = next(c for c in source.extract(text) if c['symbol'] == 'BOOL')
        self.assertEqual(row.get('guard'), ['EXTRA'])

    def test_unapproved_but_internally_consistent_source_rejected(self):
        row = copy.deepcopy(self.rows[0])
        # The current commit has exactly the same header blob, but is not the
        # reviewed source revision. Internal Git consistency is insufficient.
        row['source']['revision'] = validate.git('rev-parse', 'HEAD').decode().strip()
        with self.assertRaisesRegex(validate.Invalid, 'approved source baseline'):
            validate.validate([row])

    def test_curated_description_and_edna_survive_regeneration(self):
        rows = copy.deepcopy(self.rows)
        row = next(r for r in rows if r['symbol'] == 'xlHpcRetSuccess')
        row['description'] = 'Reviewed prose that must survive regeneration.'
        row['edna_ids'] = ['EDNA-02', 'EDNA-10']
        result = populate.build(rows, source.extract(self.text), {})
        actual = next(r for r in result if r['symbol'] == row['symbol'])
        self.assertEqual(actual['description'], row['description'])
        self.assertEqual(actual['edna_ids'], ['EDNA-02', 'EDNA-10'])

    def test_missing_approved_git_source_is_unverified(self):
        out = StringIO()
        with patch.object(validate, 'git', side_effect=validate.Invalid('missing object')), redirect_stdout(out), redirect_stderr(out):
            code = validate.main(['check'])
        self.assertEqual(code, 3)
        self.assertIn('approved source baseline unavailable', out.getvalue())
        self.assertNotIn('PASS', out.getvalue())

    def test_wrong_macro_kind_signature_pointer_and_subfamily(self):
        for symbol, field, value in [('xlflowHalt', 'kind', 'worksheet-function-id'),
                                     ('Excel4', 'signature', 'int pascal Excel4(void)'),
                                     ('PXL_HPC_ASYNC_CALLBACK', 'signature', 'void (*PXL_HPC_ASYNC_CALLBACK)(void)'),
                                     ('LPXLREF', 'target', 'XLREF12'),
                                     ('xlHpcRetSuccess', 'subfamily', 'edit-mode')]:
            row = copy.deepcopy(next(r for r in self.rows if r['symbol'] == symbol))
            row[field] = value
            with self.subTest(symbol=symbol), self.assertRaises(validate.Invalid):
                validate.validate([row])

    def test_census_metadata_cannot_contradict_its_entries(self):
        census = source.census(self.text)
        census['source']['blob_sha'] = '0' * 40
        messages, bad = validate.compare_census(self.rows, census)
        self.assertTrue(bad)
        self.assertIn('SOURCE PROVENANCE', str(messages))
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'census.json'
            for claimed, expected in ((census, 2), (source.census(self.text), 3)):
                claimed['reviewed'] = True
                claimed['reviewer'] = 'caller-supplied name cannot authenticate review'
                path.write_text(json.dumps(claimed), encoding='utf8')
                out = StringIO()
                with redirect_stdout(out), redirect_stderr(out):
                    code = validate.main(['reconcile', '--census', str(path)])
                self.assertEqual(code, expected, out.getvalue())
                self.assertIn('SOURCE COMPLETENESS NOT VERIFIED', out.getvalue())

    def test_python_only_numeric_syntax_is_rejected(self):
        for expression in ('1_000', '0b10', '0o10'):
            with self.subTest(expression=expression), self.assertRaises(source.ExtractionError):
                source.value_of(expression, {})

    def test_multiline_supported_declaration_spans(self):
        text = self.text.replace('typedef INT32 BOOL;', 'typedef\nINT32\nBOOL;', 1)
        row = next(c for c in source.extract(text) if c['symbol'] == 'BOOL')
        self.assertEqual((row['line_start'], row['line_end']), (24, 26))


if __name__ == '__main__':
    unittest.main()
