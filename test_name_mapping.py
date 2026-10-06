import unittest
from unittest.mock import patch

import pandas as pd

from report_engine import ParsedShift, _resolve_item_names, build_report


def master():
    return pd.DataFrame([{
        "Item Name Key": "ENV MASTER NAME",
        "Printing Item ID": "MJP-1",
        "Printing Item Name": "ENV MASTER NAME",
        "Machine Line Code": "L1",
        "Machine Line Name": "Line 1",
        "PCS per KG": 100,
        "Roll Weight (KG)": 0,
        "Core/Tare Weight (KG)": 0,
        "Allowance": 0,
    }])


def shift(label, name, pcs=100):
    return ParsedShift(label, label, None, pd.DataFrame([{
        "Contractor Name": "BIPIN",
        "Source Item Name": name,
        "Printing Item ID": f"DAILY-{label}",
        "Request PCS": pcs,
    }]))


class NameMappingTests(unittest.TestCase):
    def test_direct_master_match_takes_precedence_over_aliases(self):
        names = pd.DataFrame([{"m4_item_name": "ENV MASTER NAME", "mjp_item_name": "OTHER"}])
        resolved, ambiguous = _resolve_item_names((shift("A", "ENV MASTER NAME"), shift("B", "ENV MASTER NAME")), master(), names)
        self.assertEqual(resolved["ENV MASTER NAME"], "ENV MASTER NAME")
        self.assertFalse(ambiguous)

    def test_ambiguous_name_has_no_resolution(self):
        names = pd.DataFrame([
            {"m4_item_name": "ENV M4 NAME", "mjp_item_name": "ENV MASTER NAME"},
            {"m4_item_name": "ENV M4 NAME", "mjp_item_name": "ANOTHER NAME"},
        ])
        resolved, ambiguous = _resolve_item_names((shift("A", "ENV M4 NAME"), shift("B", "ENV M4 NAME")), master(), names)
        self.assertNotIn("ENV M4 NAME", resolved)
        self.assertEqual(ambiguous, {"ENV M4 NAME"})

    def test_alias_merges_shifts_before_calculation(self):
        names = pd.DataFrame([{"m4_item_name": "env m4   name", "mjp_item_name": "ENV MASTER NAME"}])
        with patch("report_engine.read_master", return_value=(master(), [], [])), patch(
            "report_engine.parse_shift_workbook",
            side_effect=[shift("A", "ENV M4 NAME"), shift("B", "ENV MASTER NAME")],
        ):
            result = build_report(None, None, None, 0, names)
        self.assertTrue(result.valid, result.errors)
        self.assertEqual(len(result.report), 1)
        self.assertEqual(result.report.loc[0, "Total PCS"], 200)
        self.assertEqual(result.report.loc[0, "Total KG"], 2)
        self.assertEqual(result.source_stats["Mapped source names"], 1)

    def test_missing_and_bad_target_block_download(self):
        for names in (None, pd.DataFrame([{"m4_item_name": "ENV M4 NAME", "mjp_item_name": "NOT IN MASTER"}])):
            with self.subTest(names=names), patch("report_engine.read_master", return_value=(master(), [], [])), patch(
                "report_engine.parse_shift_workbook",
                side_effect=[shift("A", "ENV M4 NAME"), shift("B", "ENV M4 NAME")],
            ):
                result = build_report(None, None, None, 0, names)
                self.assertFalse(result.valid)
                self.assertTrue(any("missing from master" in error for error in result.errors))


if __name__ == "__main__":
    unittest.main()
