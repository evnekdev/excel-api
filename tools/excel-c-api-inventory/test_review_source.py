"""Negative controls for the independent review, without production imports."""
import copy
import json
import unittest

import review_source as review


class IndependentReviewTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.text = review.read_source()
        cls.catalogue = [json.loads(line) for line in
                         (review.ROOT / "knowledge/excel-c-api/symbols.jsonl").read_text(encoding="utf8").splitlines()]

    def test_reviewed_source_identity_and_accounting(self):
        report = review.compare(self.text, self.catalogue)
        self.assertEqual(report["source_sha256"], review.SHA256)
        self.assertEqual(report["physical_lines"], 1480)
        self.assertEqual(report["expected_records"], 1132)
        self.assertEqual(report["differences"], [])
        self.assertEqual(sum(report["source_categories"].values()), 1094)
        self.assertEqual(report["source_categories"]["mouse-hit-test"], 35)

    def test_removed_extra_and_duplicate_record(self):
        for rows in (self.catalogue[1:], self.catalogue + [self.catalogue[0]],
                     self.catalogue + [dict(self.catalogue[0], symbol="NotInSource")]):
            with self.subTest(length=len(rows)):
                self.assertTrue(review.compare(self.text, rows)["differences"])

    def test_record_detail_mutations_are_detected(self):
        mutations = [
            ("xltypeBigData", "value", 2),
            ("xltypeBigData", "expression", "xltypeStr"),
            ("xlModeReady", "kind", "special-function-id"),
            ("BOOL", "underlying", "WORD"),
            ("LPXLOPER", "target", "XLOPER12"),
            ("xloper", "target", "XLOPER12"),
            ("XLOPER", "conditional_fields", []),
            ("PXL_HPC_ASYNC_CALLBACK", "signature", "int(void)"),
            ("Excel4", "signature", "int Excel4(void)"),
            ("Excel4", "linkage", "C++"),
            ("xlModeReady", "alias_of", "xlerrNull"),
            ("BOOL", "guard", ["__cplusplus"]),
        ]
        for symbol, field, value in mutations:
            with self.subTest(symbol=symbol, field=field):
                rows = copy.deepcopy(self.catalogue)
                next(r for r in rows if r["symbol"] == symbol)[field] = value
                self.assertTrue(review.compare(self.text, rows)["differences"])

    def test_wrong_source_anchor(self):
        for field, value in [("line_start", 1), ("line_end", 1),
                             ("blob_sha", "0" * 40), ("path", "other.h"),
                             ("revision", "HEAD"), ("repository", "other/repo")]:
            rows = copy.deepcopy(self.catalogue)
            rows[0]["source"][field] = value
            self.assertTrue(review.compare(self.text, rows)["differences"])

    def test_unsupported_source_is_not_silently_ignored(self):
        for suffix in ("\nenum Unreviewed { A };\n", "\n#undef xltypeNum\n",
                       "\n#pragma pack(1)\n", "\n#define F(x) x\n"):
            with self.subTest(suffix=suffix):
                with self.assertRaises(review.ReviewError):
                    review.expectations(self.text + suffix)

    def test_integer_grammar(self):
        self.assertEqual(review.integer("(16 | flag)", {"flag": 16384}), 16400)
        self.assertEqual(review.integer("-2", {}), -2)
        for expression in ("1 + 2", "1 << 3", "unknown", "(2", "2)", "f()", "1.5", "010", "00"):
            with self.subTest(expression=expression):
                with self.assertRaises(review.ReviewError):
                    review.integer(expression, {})

    def test_parenthesized_alias_and_forward_reference(self):
        text = "/* Error codes */\n#define A 7\n#define B (( A ))\n"
        entries, _, _ = review.expectations(text)
        self.assertEqual(entries[1]["alias_of"], "A")
        self.assertEqual(entries[1]["value"], 7)
        with self.assertRaises(review.ReviewError):
            review.expectations("/* Error codes */\n#define B A\n#define A 7\n")

    def test_scalar_guard_and_type_source_changes(self):
        entries, _, _ = review.expectations("#ifdef FEATURE\ntypedef WORD Alias;\n#endif\n")
        self.assertEqual(entries[0]["guard"], ["FEATURE"])
        self.assertEqual(entries[0]["underlying"], "WORD")
        changed = self.text.replace("typedef INT32 BOOL;", "typedef WORD BOOL;")
        self.assertTrue(review.compare(changed, self.catalogue)["differences"])

    def test_subfamilies_are_source_derived(self):
        expected, _, _ = review.expectations(self.text)
        by_name = {r["symbol"]: r for r in expected}
        for symbol, family in [("xlHpcRetSuccess", "cluster-return-code"),
                               ("xlModeReady", "edit-mode"), ("dtNil", "document-type"),
                               ("htNone", "mouse-hit-test"), ("xlUDF", "user-defined-function-id")]:
            self.assertEqual(by_name[symbol]["subfamily"], family)
        rows = copy.deepcopy(self.catalogue)
        for row in rows:
            wanted = by_name[row["symbol"]]
            if "subfamily" in wanted:
                row["subfamily"] = wanted["subfamily"]
        self.assertEqual(review.compare(self.text, rows, True)["differences"], [])
        next(r for r in rows if r["symbol"] == "htNone")["subfamily"] = "edit-mode"
        self.assertTrue(review.compare(self.text, rows, True)["differences"])


if __name__ == "__main__":
    unittest.main()
