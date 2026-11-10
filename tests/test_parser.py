import unittest
from collections import Counter
from decimal import Decimal
from pathlib import Path

from app.parser import ParseError, parse_osv

FIXTURE = Path(__file__).parent / "fixtures" / "25_ОСВ.xlsx"


class OsvParserTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.report = parse_osv(FIXTURE.read_bytes(), FIXTURE.name)

    def test_meta(self):
        self.assertEqual(self.report.account_code, "25")
        self.assertEqual(self.report.period, "3 квартал 2026 г.")
        self.assertIn("АЛАБУГА ДЕВЕЛОПМЕНТ", self.report.company)

    def test_row_count_and_types(self):
        self.assertEqual(len(self.report.rows), 158)
        kinds = Counter(row.node_type for row in self.report.rows)
        self.assertEqual(
            dict(kinds),
            {"account": 1, "item": 49, "month": 107, "total": 1},
        )

    def test_totals_match_month_details(self):
        self.assertEqual(self.report.totals[2], Decimal("109984899.49"))
        self.assertEqual(self.report.totals[3], Decimal("109984899.49"))
        month_sum = sum(
            (
                row.amounts[2]
                for row in self.report.rows
                if row.node_type == "month"
            ),
            Decimal("0.00"),
        )
        self.assertEqual(month_sum, self.report.totals[2])

    def test_tree(self):
        rows = self.report.rows
        self.assertEqual(rows[0].node_type, "account")
        self.assertIsNone(rows[0].parent_index)

        level_one = [row for row in rows if row.level == 1]
        self.assertTrue(level_one)
        self.assertTrue(all(row.parent_index == 0 for row in level_one))

        month_rows = [row for row in rows if row.node_type == "month"]
        self.assertTrue(all(row.level == 3 for row in month_rows))
        for row in month_rows:
            parent = rows[row.parent_index]
            self.assertGreaterEqual(parent.level, 1)

        totals = [row for row in rows if row.node_type == "total"]
        self.assertEqual(len(totals), 1)
        self.assertEqual(totals[0].parent_index, 0)
        self.assertEqual(totals[0].level, 1)

    def test_names_are_readable(self):
        self.assertFalse(
            any(row.name.strip() == "<...>" for row in self.report.rows)
        )


class OsvParserErrorTest(unittest.TestCase):
    def test_invalid_file(self):
        with self.assertRaises(ParseError):
            parse_osv(b"not an excel file", "25_ОСВ.xlsx")


if __name__ == "__main__":
    unittest.main()
