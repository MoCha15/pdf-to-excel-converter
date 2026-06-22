"""
Unit tests for pdf_to_excel_converter.py

Covers all pure functions — no PDF files or GUI needed.
"""

import sys
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch, call
import tempfile

# ── bootstrap: suppress the _ensure_deps() call at import time ────────────────
import importlib
_real_import = importlib.import_module

def _patched_import(name, *args, **kwargs):
    if name in ("pdfplumber",):
        raise ImportError(name)
    return _real_import(name, *args, **kwargs)

# Mock heavy deps before importing the module
sys.modules.setdefault("pdfplumber", MagicMock())
sys.modules.setdefault("openpyxl", MagicMock())
sys.modules.setdefault("openpyxl.styles", MagicMock())
sys.modules.setdefault("openpyxl.utils", MagicMock())

# Patch _ensure_deps so it doesn't try to install anything
with patch("builtins.__import__", side_effect=lambda n, *a, **k: sys.modules[n] if n in sys.modules else __builtins__.__import__(n, *a, **k)):
    pass

import importlib.util, types

# Minimal stubs so the module-level openpyxl attribute access doesn't crash
_styles = types.ModuleType("openpyxl.styles")
for _name in ("Font", "PatternFill", "Alignment", "Border", "Side"):
    setattr(_styles, _name, MagicMock(return_value=MagicMock()))
sys.modules["openpyxl.styles"] = _styles

_utils = types.ModuleType("openpyxl.utils")
_utils.get_column_letter = MagicMock(return_value="A")
sys.modules["openpyxl.utils"] = _utils

_openpyxl = types.ModuleType("openpyxl")
_openpyxl.Workbook = MagicMock()
sys.modules["openpyxl"] = _openpyxl

_pdfplumber = types.ModuleType("pdfplumber")
_pdfplumber.open = MagicMock()
sys.modules["pdfplumber"] = _pdfplumber

# Patch _ensure_deps to no-op before import
with patch.dict("sys.modules"):
    import pdf_to_excel_converter as conv


# ── helpers ────────────────────────────────────────────────────────────────────

def _word(text, x0, x1, top=10.0):
    return {"text": text, "x0": float(x0), "x1": float(x1), "top": float(top)}


# ══════════════════════════════════════════════════════════════════════════════
# _clean_row
# ══════════════════════════════════════════════════════════════════════════════

class TestCleanRow(unittest.TestCase):

    def test_none_becomes_empty_string(self):
        self.assertEqual(conv._clean_row([None, "foo", None]), ["", "foo", ""])

    def test_whitespace_normalised(self):
        self.assertEqual(conv._clean_row(["  hello   world  "]), ["hello world"])

    def test_numeric_values_stringified(self):
        self.assertEqual(conv._clean_row([42, 3.14]), ["42", "3.14"])

    def test_empty_row(self):
        self.assertEqual(conv._clean_row([]), [])

    def test_mixed_types(self):
        result = conv._clean_row([None, "  a  b  ", 0])
        self.assertEqual(result, ["", "a b", "0"])


# ══════════════════════════════════════════════════════════════════════════════
# _score_tables
# ══════════════════════════════════════════════════════════════════════════════

class TestScoreTables(unittest.TestCase):

    def test_empty_list(self):
        self.assertEqual(conv._score_tables([]), 0)

    def test_header_only_table_scores_zero(self):
        self.assertEqual(conv._score_tables([["header"]]), 0)

    def test_single_table_two_data_rows(self):
        # _score_tables receives a list-of-tables; each table is a list-of-rows.
        # A table with 3 rows scores len(table)-1 = 2.
        table = [["h1", "h2"], ["a", "b"], ["c", "d"]]
        self.assertEqual(conv._score_tables([table]), 2)

    def test_multiple_tables(self):
        t1 = [["h"], ["r1"], ["r2"]]   # 2 data rows
        t2 = [["h"], ["r1"]]           # 1 data row
        self.assertEqual(conv._score_tables([t1, t2]), 3)

    def test_empty_table_ignored(self):
        self.assertEqual(conv._score_tables([[], ["h", "r"]]), 1)


# ══════════════════════════════════════════════════════════════════════════════
# _detect_columns
# ══════════════════════════════════════════════════════════════════════════════

class TestDetectColumns(unittest.TestCase):

    def test_empty_returns_empty(self):
        anchors, boundaries = conv._detect_columns([])
        self.assertEqual(anchors, [])
        self.assertEqual(boundaries, [])

    def test_single_word(self):
        anchors, boundaries = conv._detect_columns([_word("Date", 10, 30)])
        self.assertEqual(len(anchors), 1)
        self.assertEqual(boundaries, [])

    def test_two_well_separated_words_form_two_columns(self):
        # gap of 50 units — clearly two separate columns
        words = [_word("Date", 0, 30), _word("Amount", 80, 120)]
        anchors, boundaries = conv._detect_columns(words)
        self.assertEqual(len(anchors), 2)
        self.assertEqual(len(boundaries), 1)
        self.assertAlmostEqual(boundaries[0], 55.0)   # midpoint of 30 and 80

    def test_close_words_cluster_into_one_column(self):
        # "Chq" and "No" are 2 units apart — same column label
        words = [_word("Chq", 0, 20), _word("No", 22, 40), _word("Balance", 100, 150)]
        anchors, boundaries = conv._detect_columns(words)
        self.assertEqual(len(anchors), 2)   # "Chq No" + "Balance"

    def test_boundary_is_midpoint_between_clusters(self):
        words = [_word("A", 0, 10), _word("B", 60, 80)]
        _, boundaries = conv._detect_columns(words)
        self.assertAlmostEqual(boundaries[0], 35.0)   # (10 + 60) / 2


# ══════════════════════════════════════════════════════════════════════════════
# _looks_like_data_header
# ══════════════════════════════════════════════════════════════════════════════

class TestLooksLikeDataHeader(unittest.TestCase):

    def test_empty_row(self):
        self.assertFalse(conv._looks_like_data_header([]))

    def test_single_keyword_not_enough(self):
        # "Date" is a keyword but "Title" is not — only 1 keyword match, need ≥1
        # but also need ≥2 non-empty cells (passes) AND a keyword hit.
        # "Date" alone IS a keyword so this returns True — test the truly-no-keyword case.
        self.assertFalse(conv._looks_like_data_header(["Invoice", "Title"]))

    def test_two_keywords_returns_true(self):
        self.assertTrue(conv._looks_like_data_header(["Date", "Balance", "Description"]))

    def test_case_insensitive(self):
        self.assertTrue(conv._looks_like_data_header(["DATE", "DEBIT", "CREDIT"]))

    def test_non_financial_row(self):
        self.assertFalse(conv._looks_like_data_header(["Company Name", "Address"]))

    def test_single_cell_row(self):
        self.assertFalse(conv._looks_like_data_header(["Date"]))

    def test_all_empty_cells(self):
        self.assertFalse(conv._looks_like_data_header(["", "  ", ""]))


# ══════════════════════════════════════════════════════════════════════════════
# _norm_cell
# ══════════════════════════════════════════════════════════════════════════════

class TestNormCell(unittest.TestCase):

    def test_empty_string(self):
        self.assertEqual(conv._norm_cell(""), "")

    def test_single_word(self):
        self.assertEqual(conv._norm_cell("Date"), "date")

    def test_multi_word_returns_first_token(self):
        self.assertEqual(conv._norm_cell("Particulars otil"), "particulars")

    def test_strips_whitespace(self):
        self.assertEqual(conv._norm_cell("  Balance  "), "balance")


# ══════════════════════════════════════════════════════════════════════════════
# _all_headers_match
# ══════════════════════════════════════════════════════════════════════════════

class TestAllHeadersMatch(unittest.TestCase):

    def _table(self, header, *data_rows):
        return {"rows": [header] + list(data_rows)}

    def test_single_table_returns_false(self):
        t = self._table(["Date", "Balance"], ["01 Jan", "100"])
        self.assertFalse(conv._all_headers_match([t]))

    def test_two_matching_headers(self):
        hdr = ["Date", "Debit", "Credit", "Balance"]
        t1 = self._table(hdr, ["01 Jan", "10", "", "90"])
        t2 = self._table(hdr, ["02 Jan", "", "20", "110"])
        self.assertTrue(conv._all_headers_match([t1, t2]))

    def test_mismatched_headers(self):
        t1 = self._table(["Date", "Debit", "Balance"], ["01 Jan", "10", "90"])
        t2 = self._table(["Date", "Credit", "Balance"], ["02 Jan", "20", "110"])
        self.assertFalse(conv._all_headers_match([t1, t2]))

    def test_non_data_tables_ignored(self):
        # table with non-financial header should not count toward the 2-real-tables minimum
        cover = self._table(["Company Name", "Report 2024"])
        data  = self._table(["Date", "Balance"], ["01 Jan", "100"])
        self.assertFalse(conv._all_headers_match([cover, data]))

    def test_bleed_text_in_header_normalised(self):
        # "Particulars otil" normalises to "particulars" — should still match
        t1 = self._table(["Date", "Particulars otil", "Balance"], ["r1", "r2", "r3"])
        t2 = self._table(["Date", "Particulars",      "Balance"], ["r1", "r2", "r3"])
        self.assertTrue(conv._all_headers_match([t1, t2]))


# ══════════════════════════════════════════════════════════════════════════════
# _combine_tables
# ══════════════════════════════════════════════════════════════════════════════

class TestCombineTables(unittest.TestCase):

    def _table(self, label, header, *data_rows):
        return {"label": label, "rows": [header] + list(data_rows)}

    def test_two_data_tables_merged(self):
        hdr = ["Date", "Debit", "Balance"]
        t1 = self._table("Page1", hdr, ["01 Jan", "10", "90"])
        t2 = self._table("Page2", hdr, ["02 Jan", "20", "70"])
        result = conv._combine_tables([t1, t2])
        combined = next(t for t in result if t["label"] == "Combined")
        self.assertEqual(combined["rows"][0], hdr)
        self.assertEqual(len(combined["rows"]), 3)   # header + 2 data rows

    def test_non_data_table_kept_separate(self):
        cover = self._table("Cover", ["Company", "Address"], ["Acme", "123 St"])
        hdr   = ["Date", "Balance"]
        t1    = self._table("Page1", hdr, ["01 Jan", "100"])
        t2    = self._table("Page2", hdr, ["02 Jan", "200"])
        result = conv._combine_tables([cover, t1, t2])
        labels = [t["label"] for t in result]
        self.assertIn("Cover", labels)
        self.assertIn("Combined", labels)

    def test_header_not_duplicated(self):
        hdr = ["Date", "Credit", "Balance"]
        t1 = self._table("P1", hdr, ["d1", "c1", "b1"])
        t2 = self._table("P2", hdr, ["d2", "c2", "b2"])
        result = conv._combine_tables([t1, t2])
        combined = next(t for t in result if t["label"] == "Combined")
        # first row is the header; subsequent rows are data only
        self.assertEqual(combined["rows"][0], hdr)
        self.assertNotEqual(combined["rows"][1], hdr)


# ══════════════════════════════════════════════════════════════════════════════
# _output_path
# ══════════════════════════════════════════════════════════════════════════════

class TestOutputPath(unittest.TestCase):

    def test_basic_path(self):
        with tempfile.TemporaryDirectory() as d:
            pdf = str(Path(d) / "report.pdf")
            out = conv._output_path(pdf)
            self.assertEqual(Path(out).name, "report_tables.xlsx")
            self.assertEqual(Path(out).parent, Path(d))

    def test_increments_when_file_exists(self):
        with tempfile.TemporaryDirectory() as d:
            # pre-create the first candidate
            (Path(d) / "report_tables.xlsx").touch()
            pdf = str(Path(d) / "report.pdf")
            out = conv._output_path(pdf)
            self.assertEqual(Path(out).name, "report_tables_1.xlsx")

    def test_increments_twice(self):
        with tempfile.TemporaryDirectory() as d:
            (Path(d) / "report_tables.xlsx").touch()
            (Path(d) / "report_tables_1.xlsx").touch()
            pdf = str(Path(d) / "report.pdf")
            out = conv._output_path(pdf)
            self.assertEqual(Path(out).name, "report_tables_2.xlsx")


if __name__ == "__main__":
    unittest.main()
