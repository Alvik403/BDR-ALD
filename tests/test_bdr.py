import unittest
from decimal import Decimal
from types import SimpleNamespace

from app.bdr import build_bdr


def row(
    account,
    target,
    article,
    month,
    debit="0",
    credit="0",
    status="mapped",
    kind="nomenclature",
):
    return SimpleNamespace(
        account_code=account,
        target=target,
        article=article,
        month=month,
        debit=Decimal(debit),
        credit=Decimal(credit),
        status=status,
        mapping_kind=kind,
    )


class BdrBuildTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.rows = []
        for month in ("Июль 26", "Август 26", "Сентябрь 26"):
            cls.rows.extend(
                [
                    row("90.01", "ГП", "", month, credit="100"),
                    row("90.01", "ПИР", "", month, credit="50"),
                    row("90.01", "Прочее", "", month, credit="10"),
                    row("90.02", "СМР", "", month, debit="60"),
                    row("90.02", "ПИР", "", month, debit="20"),
                    row("90.02", "Прочее", "", month, debit="5"),
                    row(
                        "90.02",
                        "Строители (КУП)",
                        "Оплата труда",
                        month,
                        debit="30",
                        kind="department",
                    ),
                    row("26", "Общехозяйственные расходы", "Оплата труда", month, debit="10", kind="constant"),
                    row("26", "Общехозяйственные расходы", "Аренда", month, debit="7", kind="constant"),
                    row("25", "Общепроизводственные расходы", "Оплата труда", month, debit="4", kind="constant"),
                    row("25", "Общепроизводственные расходы", "Аренда", month, debit="3", kind="constant"),
                    row("91.01", "% к получению", "", month, credit="2"),
                    row("91.01", "Прочие доходы", "", month, credit="8"),
                    row("91.02", "% к уплате", "", month, debit="1"),
                    row("91.02", "% за ЗУ", "", month, debit="1"),
                    row("91.02", "Прочие расходы", "", month, debit="6"),
                ]
            )
        cls.rows.append(
            row(
                "90.01",
                "ГП",
                "",
                "Июль 26",
                credit="999",
                status="unmapped",
            )
        )
        cls.bdr = build_bdr(cls.rows, {7, 8, 9})
        cls.by_label = {
            (item.section, item.param): item for item in cls.bdr["rows"]
        }

    def test_summary(self):
        summary = self.bdr["summary"]
        self.assertEqual(summary["revenue"], Decimal("480.00"))
        self.assertEqual(summary["cost"], Decimal("255.00"))
        self.assertEqual(summary["gross"], Decimal("225.00"))
        self.assertEqual(summary["fot"], Decimal("120.00"))
        self.assertEqual(summary["profit_sales"], Decimal("84.00"))
        self.assertEqual(summary["profit_before_tax"], Decimal("90.00"))

    def test_25_is_shown_but_excluded(self):
        opex25 = self.by_label[("Расходы", "Общепроизводственные расходы")]
        opex26 = self.by_label[("Расходы", "Общехозяйственные расходы")]
        self.assertEqual(opex25.values["q3"], Decimal("9.00"))
        self.assertEqual(opex26.values["q3"], Decimal("21.00"))
        self.assertEqual(
            self.bdr["summary"]["profit_sales"], Decimal("84.00")
        )

    def test_rows_and_periods(self):
        revenue = self.by_label[("Выручка", "")]
        self.assertEqual(revenue.values["q3"], Decimal("480.00"))
        self.assertEqual(revenue.values["m7"], Decimal("160.00"))
        self.assertIsNone(revenue.values["q1"])
        self.assertIsNone(revenue.values["nine"])
        self.assertIsNone(revenue.values["y2026"])

        profit = self.by_label[("Прибыль до НО", "БУ")]
        self.assertEqual(profit.values["q3"], Decimal("90.00"))

    def test_unmapped_does_not_affect_totals(self):
        self.assertTrue(self.bdr["has_unmapped"])
        self.assertEqual(self.bdr["summary"]["revenue"], Decimal("480.00"))

    def test_unmapped_fot_is_still_included(self):
        rows = [
            row(
                "90.02",
                "",
                "Оплата труда",
                "Июль 26",
                debit="10",
                status="unmapped",
                kind="department",
            )
        ]
        bdr = build_bdr(rows, {7, 8, 9})
        self.assertEqual(bdr["summary"]["fot"], Decimal("10.00"))
        self.assertFalse(bdr["has_unmapped"])


if __name__ == "__main__":
    unittest.main()
