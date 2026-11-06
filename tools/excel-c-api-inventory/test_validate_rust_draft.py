"""Integrity tests for the bounded draft assessment."""
import copy
import contextlib
import io
import json
import subprocess
import sys
import types
import unittest
from unittest import mock

import validate_rust_draft as audit
import summarize_rust_draft as summary


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

    def test_duplicate_json_keys_are_rejected(self):
        with self.assertRaises(ValueError):
            audit.catalogue_validator.read_jsonl(b'{"capi_id":"CAPI-000001","capi_id":"CAPI-000002"}\n')

    def test_unknown_schema_version_is_rejected(self):
        rows = copy.deepcopy(self.rows)
        rows[0]['assessment_version'] = 3
        self.assert_invalid(rows)
        rows = copy.deepcopy(self.rows)
        rows[0]['rust_revision'] = '0' * 40
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
        row = next(row for row in rows if row['functional_state'] == 'typed-operation-present')
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

    def test_abi_evidence_must_match_checker_and_target(self):
        rows = copy.deepcopy(self.rows)
        row = next(row for row in rows if row['abi_evidence'])
        row['abi_evidence']['target'] = 'i686-pc-windows-msvc'
        self.assert_invalid(rows)
        rows = copy.deepcopy(self.rows)
        row = next(row for row in rows if row['abi_evidence'])
        row['abi_evidence']['line'] = 999999
        self.assert_invalid(rows)
        rows = copy.deepcopy(self.rows)
        row = next(row for row in rows if row['test_state'] == 'abi-check-passed-this-run')
        row['abi_evidence'] = None
        self.assert_invalid(rows)

    def test_unsupported_decoder_and_event_mappings_cannot_be_relabelled(self):
        rows = copy.deepcopy(self.rows)
        row = next(row for row in rows if row['functional_state'] == 'decoding-explicitly-unsupported')
        row['functional_state'] = 'decode-branch-present'
        self.assert_invalid(rows)
        rows = copy.deepcopy(self.rows)
        row = next(row for row in rows if row['functional_state'] == 'event-registration-present')
        row['functional_evidence']['configuration'] = 'default'
        self.assert_invalid(rows)

    def test_descriptor_absence_cannot_be_forged(self):
        rows = copy.deepcopy(self.rows)
        row = next(row for row in rows if row['descriptor_observation'] == 'present')
        row['descriptor_observation'] = 'not-observed'
        self.assert_invalid(rows)

    def test_related_materialization_reference_must_match(self):
        rows = copy.deepcopy(self.rows)
        row = next(row for row in rows if row['related_evidence'])
        row['related_evidence'][0]['line'] = 1
        self.assert_invalid(rows)
        rows = copy.deepcopy(self.rows)
        row = next(row for row in rows if len(row['related_evidence']) == 1)
        row['related_evidence'].append(copy.deepcopy(row['related_evidence'][0]))
        self.assert_invalid(rows)

    def test_fp12_registration_observation_requires_the_right_enum(self):
        rows = copy.deepcopy(self.rows)
        row = next(row for row in rows if row['functional_state'] == 'fp12-registration-type-not-observed')
        row['functional_evidence']['symbol'] = 'ExcelReturnType'
        self.assert_invalid(rows)

    def test_core_occurrence_counts_are_checked(self):
        rows = copy.deepcopy(self.rows)
        rows[0]['core_name_occurrences'] += 1
        self.assert_invalid(rows)

    def test_supplemental_mention_locations_are_checked(self):
        rows = copy.deepcopy(self.rows)
        row = next(row for row in rows if row['supplemental_mentions']['first_three'])
        row['supplemental_mentions']['first_three'][0]['line'] = 999999
        self.assert_invalid(rows)

    def test_penhelper_only_selector_cannot_be_claimed_as_generic(self):
        rows = copy.deepcopy(self.rows)
        row = next(row for row in rows if row['dispatch_relation'] == 'lpenhelper-selector-only')
        row['dispatch_relation'] = 'generic-selector-availability-unassessed'
        self.assert_invalid(rows)

    def test_type_target_must_resolve_to_the_source_aggregate(self):
        rows = copy.deepcopy(self.rows)
        row = next(row for row in rows if row['type_target_capi_id'])
        row['type_target_capi_id'] = 'CAPI-999999'
        self.assert_invalid(rows)

    def test_edna_relation_cannot_be_fabricated(self):
        rows = copy.deepcopy(self.rows)
        row = next(row for row in rows if row['edna_relations'])
        row['edna_relations'][0]['id'] = 'EDNA-99'
        self.assert_invalid(rows)

    def test_generated_thunk_compile_evidence_must_match(self):
        rows = copy.deepcopy(self.rows)
        row = next(row for row in rows if row['test_state'] == 'compile-check-passed-this-run')
        row['test_evidence']['line'] = 1
        self.assert_invalid(rows)

    def test_legacy_absence_cannot_be_assigned_to_an_unrelated_symbol(self):
        rows = copy.deepcopy(self.rows)
        row = next(row for row in rows if row['raw_state'] == 'numeric-value-matched')
        row['raw_state'] = 'legacy-public-declaration-not-observed'
        self.assert_invalid(rows)

    def test_changed_source_is_reported_stale(self):
        header = (audit.ROOT / 'tools/Excel2013XLLSDK/INCLUDE/XLCALL.H').read_bytes().replace(b'\r\n', b'\n')
        with mock.patch.object(audit.subprocess, 'check_output', return_value=header), mock.patch.object(audit.subprocess, 'run', return_value=types.SimpleNamespace(returncode=1)):
            with self.assertRaisesRegex(ValueError, 'reassessment required'):
                audit.verify_current_inputs()

    def test_missing_catalogue_history_is_not_verified(self):
        missing = audit.catalogue_validator.HistoryUnavailable('missing baseline')
        with mock.patch.object(audit, 'verify_current_inputs', side_effect=missing), contextlib.redirect_stderr(io.StringIO()) as output:
            self.assertEqual(audit.main(), 3)
        self.assertIn('NOT VERIFIED', output.getvalue())

    def test_missing_pinned_sdk_object_is_not_verified(self):
        error = subprocess.CalledProcessError(128, ['git', 'show'])
        with mock.patch.object(audit.subprocess, 'check_output', side_effect=error), contextlib.redirect_stderr(io.StringIO()) as output:
            self.assertEqual(audit.main(), 3)
        self.assertIn('pinned SDK source object unavailable', output.getvalue())

    def test_documented_summary_matches_assessment(self):
        command = [sys.executable, str(audit.ROOT / 'tools/excel-c-api-inventory/summarize_rust_draft.py')]
        table = subprocess.check_output(command, cwd=audit.ROOT, text=True, encoding='utf8')
        document = (audit.ROOT / 'knowledge/excel-c-api/rust-assessment-draft.md').read_text(encoding='utf8')
        self.assertIn(table, document)

    def test_source_only_test_is_not_counted_as_passed(self):
        self.assertEqual(summary.observed_count([{'test_state': 'test-source-present'}], 'test_state'), 0)

    def test_indented_platform_guard_is_preserved(self):
        rows = copy.deepcopy(self.rows)
        rows[0]['test_evidence']['configuration'] = 'windows-x86_64'
        self.assertEqual(audit.verify(rows), 1132)
        rows[0]['test_evidence']['configuration'] = 'default'
        self.assert_invalid(rows)

    def test_comment_and_import_are_not_return_materialization(self):
        for line in (5, 12):
            rows = copy.deepcopy(self.rows)
            row = next(row for row in rows if row['capi_id'] == 'CAPI-000049')
            row['related_evidence'][0]['line'] = line
            self.assert_invalid(rows)

    def test_build_configuration_changes_require_reassessment(self):
        original_run = audit.subprocess.run
        for changed_path in ('Cargo.toml', 'Cargo.lock', 'crates/excel-api/Cargo.toml',
                             'crates/excel-api-macros/build.rs', '.cargo/config.toml'):
            def changed_diff(command, **kwargs):
                if command[:3] == ['git', 'diff', '--quiet']:
                    scopes = command[command.index('--') + 1:]
                    covered = any(changed_path == path or changed_path.startswith(path + '/') for path in scopes)
                    return types.SimpleNamespace(returncode=int(covered))
                return original_run(command, **kwargs)
            with self.subTest(path=changed_path), mock.patch.object(audit.subprocess, 'run', side_effect=changed_diff):
                with self.assertRaisesRegex(ValueError, 'reassessment required'):
                    audit.verify_current_inputs()

    def test_typed_operation_needs_implementation_and_correct_version(self):
        rows = copy.deepcopy(self.rows)
        row = next(r for r in rows if r['functional_state'] == 'typed-operation-present')
        row['related_evidence'] = [e for e in row['related_evidence'] if e['role'] != 'typed-operation']
        self.assert_invalid(rows)
        rows = copy.deepcopy(self.rows)
        row = next(r for r in rows if r['functional_state'] == 'typed-operation-present')
        row['assessment_version'] = 1
        self.assert_invalid(rows)

    def test_lpenhelper_and_legacy_boundaries_do_not_apply_to_modern_calls(self):
        for field, state in (('raw_state', 'lpenhelper-declaration-not-observed'),
                             ('functional_state', 'lpenhelper-operation-not-observed'),
                             ('functional_state', 'legacy-outside-excel12-target')):
            rows = copy.deepcopy(self.rows)
            row = next(r for r in rows if r['capi_id'] == 'CAPI-000010')
            row['assessment_version'] = 2
            row[field] = state
            self.assert_invalid(rows)

    def test_generic_backend_does_not_prove_arbitrary_selector_support(self):
        rows = copy.deepcopy(self.rows)
        backend = next(r for r in rows if r['functional_state'] == 'generic-backend-present')
        unknown = next(r for r in rows if r['functional_state'] == 'unassessed' and r['dispatch_relation'])
        unknown['assessment_version'] = 2
        unknown['functional_state'] = backend['functional_state']
        unknown['functional_evidence'] = copy.deepcopy(backend['functional_evidence'])
        self.assert_invalid(rows)

    def test_command_registration_gap_is_source_backed_and_bounded(self):
        self.assertEqual(audit.command_registration_counts(), (9, 10))
        rows = copy.deepcopy(self.rows)
        row = next(r for r in rows if r['functional_state'] == 'typed-registration-partial')
        with mock.patch.object(audit, 'command_registration_counts', return_value=(10, 10)):
            self.assert_invalid(rows)
        row['functional_state'] = 'typed-operation-present'
        self.assert_invalid(rows)
        row['functional_state'] = 'typed-registration-partial'
        row['capi_id'] = 'CAPI-000010'
        self.assert_invalid(rows)

    def test_generated_alias_must_match_source_and_generator(self):
        rows = copy.deepcopy(self.rows)
        evidence = next(e for r in rows for e in r['related_evidence'] if e['role'] == 'generated-abi-alias')
        evidence['line'] = 20 if evidence['line'] == 23 else 23
        self.assert_invalid(rows)


if __name__ == '__main__':
    unittest.main()
