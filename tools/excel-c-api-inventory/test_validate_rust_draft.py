"""Integrity tests for the bounded draft assessment."""
import copy
import json
import unittest
from unittest import mock

import validate_rust_draft as audit


class DraftValidationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.rows = [json.loads(line) for line in audit.DRAFT.read_text(encoding='utf8').splitlines()]

    def test_current_draft(self):
        self.assertEqual(audit.verify(self.rows), 1132)

    def assert_invalid(self, rows):
        with self.assertRaises(ValueError):
            audit.verify(rows)

    def test_missing_duplicate_and_unknown_ids(self):
        self.assert_invalid(self.rows[:-1])
        self.assert_invalid(self.rows + [self.rows[0]])
        rows = copy.deepcopy(self.rows)
        rows[0]['capi_id'] = 'CAPI-999999'
        self.assert_invalid(rows)

    def test_claims_and_references_must_be_supported(self):
        rows = copy.deepcopy(self.rows)
        rows[0]['host_state'] = 'pass'
        self.assert_invalid(rows)
        rows = copy.deepcopy(self.rows)
        row = next(row for row in rows if row['raw_state'] == 'numeric-value-matched')
        row['exact_name_candidates'][0]['line'] = 999999
        self.assert_invalid(rows)
        rows = copy.deepcopy(self.rows)
        row = next(row for row in rows if row['functional_state'] == 'descriptor-present')
        row['functional_evidence']['symbol'] = 'XL_FAKE'
        self.assert_invalid(rows)

    def test_non_numeric_record_cannot_claim_numeric_match(self):
        rows = copy.deepcopy(self.rows)
        row = next(row for row in rows if row['raw_state'] == 'declaration-present-abi-unassessed')
        row['raw_state'] = 'numeric-value-matched'
        self.assert_invalid(rows)

    def test_incorrect_numeric_value_and_feature_are_rejected(self):
        constants = audit.numeric_constants()
        symbol = 'xltypeNum'
        constants[symbol] += 1
        with mock.patch.object(audit, 'numeric_constants', return_value=constants):
            self.assert_invalid(self.rows)
        rows = copy.deepcopy(self.rows)
        row = next(row for row in rows if row['functional_evidence'] and row['functional_evidence']['configuration'] == 'xlcontime-research')
        row['functional_evidence']['configuration'] = 'default'
        self.assert_invalid(rows)

    def test_missing_search_lead_is_rejected(self):
        rows = copy.deepcopy(self.rows)
        row = next(row for row in rows if row['exact_name_candidates'])
        row['exact_name_candidates'] = []
        self.assert_invalid(rows)


if __name__ == '__main__':
    unittest.main()
