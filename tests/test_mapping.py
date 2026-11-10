import unittest
from io import BytesIO
from types import SimpleNamespace

from openpyxl import Workbook, load_workbook

from app.mapping import (
    CONSTANT_TARGETS,
    MappingImportError,
    build_fot_mapping_template,
    build_mapping_template,
    normalize_key,
    parse_mapping_workbook,
    resolve_mapping,
)


def row(id, parent_id, node_type, name, level, month=None):
    return SimpleNamespace(
        id=id,
        parent_id=parent_id,
        node_type=node_type,
        name=name,
        level=level,
        month=month,
    )


class NormalizeKeyTest(unittest.TestCase):
    def test_case_spaces_quotes_and_yo(self):
        self.assertEqual(
            normalize_key('  "Синергия  10.1,2"  '),
            "синергия 10.1,2",
        )
        self.assertEqual(normalize_key("Взносы в ФСС от НС и ПЗ"), "взносы в фсс от нс и пз")
        self.assertEqual(normalize_key("Аренда."), "аренда")
        self.assertEqual(normalize_key("Счёт ёжик"), "счет ежик")


class ResolveMappingTest(unittest.TestCase):
    def setUp(self):
        self.maps = {
            "nomenclature": {
                ("90.01", 2026, "синергия 10.1,2"): "Коммерческая недвижимость",
                ("90.02", 2026, "синергия 10.1,2"): "Коммерческая недвижимость",
            },
            "department": {
                ("90.02", 2026, "отдел проектирования"): "ФОТ АУП",
                ("25", 2026, "отдел проектирования"): "ФОТ АУП",
            },
            "article": {
                ("91.01", 0, "благотворительность"): "Прочие доходы",
            },
        }

    def test_9001_uses_nomenclature_group(self):
        target, kind, key = resolve_mapping(
            "90.01",
            "Отдел",
            "Синергия 10.1,2",
            "",
            self.maps,
            "90.01",
            2026,
        )
        self.assertEqual(target, "Коммерческая недвижимость")
        self.assertEqual(kind, "nomenclature")
        self.assertEqual(key, "Синергия 10.1,2")

    def test_9002_labor_uses_department(self):
        target, kind, key = resolve_mapping(
            "90.02",
            "Отдел проектирования",
            "Группа",
            "Оплата труда",
            self.maps,
            "90.02",
            2026,
        )
        self.assertEqual(target, "ФОТ АУП")
        self.assertEqual(kind, "department")

    def test_9002_non_labor_uses_nomenclature(self):
        target, kind, key = resolve_mapping(
            "90.02",
            "Отдел проектирования",
            "Синергия 10.1,2",
            "Аренда",
            self.maps,
            "90.02",
            2026,
        )
        self.assertEqual(target, "Коммерческая недвижимость")
        self.assertEqual(kind, "nomenclature")

    def test_25_and_26_constants(self):
        target, kind, _ = resolve_mapping(
            "25", "Отдел", "", "Аренда", self.maps, "25", 2026
        )
        self.assertEqual(target, CONSTANT_TARGETS["25"])
        self.assertEqual(kind, "constant")
        target, kind, _ = resolve_mapping(
            "26", "Отдел", "", "Аренда", self.maps, "26", 2026
        )
        self.assertEqual(target, CONSTANT_TARGETS["26"])
        self.assertEqual(kind, "constant")

    def test_25_labor_uses_department(self):
        target, kind, _ = resolve_mapping(
            "25",
            "Отдел проектирования",
            "",
            "Оплата больничного",
            self.maps,
            "25",
            2026,
        )
        self.assertEqual(target, "ФОТ АУП")
        self.assertEqual(kind, "department")

    def test_9101_uses_article(self):
        target, kind, _ = resolve_mapping(
            "91.01",
            "Отдел",
            "",
            "Благотворительность",
            self.maps,
            "91.01",
            2026,
        )
        self.assertEqual(target, "Прочие доходы")
        self.assertEqual(kind, "article")

    def test_article_mapping_is_account_specific(self):
        maps = {
            "article": {
                ("91.01", 0, "госпошлина"): "Прочие доходы",
                ("91.02", 0, "госпошлина"): "Прочие расходы",
            }
        }
        first, _, _ = resolve_mapping(
            "91.01", "", "", "Госпошлина", maps, "91.01", 2026
        )
        second, _, _ = resolve_mapping(
            "91.02", "", "", "Госпошлина", maps, "91.02", 2026
        )
        self.assertEqual(first, "Прочие доходы")
        self.assertEqual(second, "Прочие расходы")

    def test_unmapped(self):
        target, kind, key = resolve_mapping(
            "90.01",
            "Отдел",
            "Неизвестная группа",
            "",
            self.maps,
            "90.01",
            2026,
        )
        self.assertIsNone(target)
        self.assertEqual(kind, "nomenclature")
        self.assertEqual(key, "Неизвестная группа")


class MappingImportTest(unittest.TestCase):
    def _bdr_like_workbook(self) -> bytes:
        workbook = Workbook()
        articles = workbook.active
        articles.title = "B–D Статьи"
        articles.append(
            [
                "Счет 1С",
                "Статьи затрат / Прочие доходы, расходы",
                "Признак",
            ]
        )
        articles.append(["91.02", "Аренда", "Признак 1"])
        articles.append(["91.01", "Аренда", "Признак 2"])
        articles.append(["91.02", "Без признака", ""])

        nomenclature = workbook.create_sheet("G–J Номенклатурные группы")
        nomenclature.append(
            ["Счет 1С", "Год", "Номенклатурные группы из 1С", "для БДР"]
        )
        nomenclature.append(
            ["90.01", 2026, "Синергия 10.1,2", "Коммерческая недвижимость"]
        )

        departments = workbook.create_sheet("M–Q ФОТ по подразделениям")
        departments.append(
            [
                "Счет 1С",
                "Год",
                "Признак (ТОЛЬКО ФОТ)",
                "Подразделение из 1С",
                "для БДР",
            ]
        )
        departments.append(
            ["25", 2026, "ФОТ", "Отдел проектирования", "ФОТ АУП"]
        )

        buffer = BytesIO()
        workbook.save(buffer)
        return buffer.getvalue()

    def test_parses_bdr_columns(self):
        items, skipped, conflicts = parse_mapping_workbook(
            self._bdr_like_workbook()
        )
        self.assertEqual(len(items), 8)
        self.assertEqual(skipped, 1)
        self.assertEqual(conflicts, 0)
        self.assertEqual(
            {item["kind"] for item in items},
            {"article", "nomenclature", "department"},
        )
        article_9102 = next(
            item
            for item in items
            if item["kind"] == "article"
            and item["account_code"] == "91.02"
        )
        article_9101 = next(
            item
            for item in items
            if item["kind"] == "article"
            and item["account_code"] == "91.01"
        )
        self.assertEqual(article_9102["target"], "Признак 1")
        self.assertEqual(article_9101["target"], "Признак 2")

        department = next(
            item for item in items if item["kind"] == "department"
        )
        self.assertEqual(department["payload"]["flag"], "ФОТ")
        self.assertEqual(department["payload"]["year"], 2026)

        special = next(
            item
            for item in items
            if item["kind"] == "nomenclature"
            and item["account_code"] == "90.01"
            and item["source_key"] == "сдача имущества в аренду"
        )
        self.assertEqual(special["target"], "Прочее")

    def test_parses_single_sheet_by_title(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "M–Q ФОТ по подразделениям"
        sheet.append(
            [
                "Счет 1С",
                "Год",
                "Признак (ТОЛЬКО ФОТ)",
                "Подразделение из 1С",
                "для БДР",
            ]
        )
        sheet.append(["25", 2026, "ФОТ", "Отдел проектирования", "ФОТ АУП"])
        buffer = BytesIO()
        workbook.save(buffer)

        items, skipped, conflicts = parse_mapping_workbook(buffer.getvalue())
        self.assertEqual(skipped, 0)
        self.assertEqual(conflicts, 0)
        self.assertEqual(len(items), 1)
        self.assertEqual(items[0]["kind"], "department")

    def test_empty_file_raises(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "B–D Статьи"
        sheet.append(
            [
                "Счет 1С",
                "Статьи затрат / Прочие доходы, расходы",
                "Признак",
            ]
        )
        sheet.append(["91.02", "Аренда", ""])
        buffer = BytesIO()
        workbook.save(buffer)
        with self.assertRaises(MappingImportError):
            parse_mapping_workbook(buffer.getvalue())

    def test_parses_parallel_layout(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "мэппинг"
        headers = {
            2: "Счет 1С",
            3: "Статьи затрат / Прочие доходы, расходы",
            4: "Признак",
            7: "Счет 1С",
            8: "Год",
            9: "Номенклатурные группы из 1С",
            10: "для БДР",
            13: "Счет 1С",
            14: "Год",
            15: "Признак (ТОЛЬКО ФОТ)",
            16: "Подразделение из 1С",
            17: "для БДР",
        }
        for col, value in headers.items():
            sheet.cell(row=1, column=col, value=value)

        rows = [
            {
                2: "91.01",
                3: "Госпошлина",
                4: "Прочие доходы",
                7: "90.01",
                8: 2026,
                9: "Группа А",
                10: "Цель А",
                13: "25",
                14: 2026,
                15: "ФОТ",
                16: "Отдел 1",
                17: "ФОТ Цель",
            },
            {
                2: "91.02",
                3: "Госпошлина",
                4: "Прочие расходы",
                7: "90.02",
                8: 2026,
                9: "Группа А",
                10: "Цель Б",
                13: "90.02",
                14: 2026,
                15: "ФОТ",
                16: "Отдел 1",
                17: "ФОТ Цель",
            },
            {
                2: "25",
                3: "Аренда",
                4: "Игнорируется",
                7: "90.01",
                8: 2025,
                9: "Старая группа",
                10: "Старая цель",
                13: "26",
                14: 2026,
                15: "ФОТ",
                16: "Отдел 2",
                17: "ФОТ Цель 2",
            },
            {
                2: "91.01",
                3: "<...>",
                4: "Прочие доходы",
            },
        ]
        for index, values in enumerate(rows, start=2):
            for col, value in values.items():
                sheet.cell(row=index, column=col, value=value)

        buffer = BytesIO()
        workbook.save(buffer)
        items, skipped, conflicts = parse_mapping_workbook(buffer.getvalue())

        self.assertEqual(skipped, 0)
        self.assertEqual(conflicts, 0)
        self.assertEqual(len(items), 11)

        article_9101 = next(
            item
            for item in items
            if item["kind"] == "article"
            and item["account_code"] == "91.01"
        )
        article_9102 = next(
            item
            for item in items
            if item["kind"] == "article"
            and item["account_code"] == "91.02"
        )
        self.assertEqual(article_9101["target"], "Прочие доходы")
        self.assertEqual(article_9102["target"], "Прочие расходы")

        self.assertFalse(
            any(
                item["kind"] == "article"
                and item["account_code"] == "25"
                for item in items
            )
        )
        self.assertFalse(
            any(
                item["kind"] == "nomenclature"
                and item["source_key"] == "старая группа"
                for item in items
            )
        )

    def test_department_rows_only_for_fot_and_supported_accounts(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "M–Q ФОТ по подразделениям"
        sheet.append(
            [
                "Счет 1С",
                "Год",
                "Признак (ТОЛЬКО ФОТ)",
                "Подразделение из 1С",
                "для БДР",
            ]
        )
        sheet.append(["25", 2026, "ФОТ", "Отдел 1", "ФОТ Цель"])
        sheet.append(["25", 2026, "Не ФОТ", "Отдел 2", "Нельзя"])
        sheet.append(["90.01", 2026, "ФОТ", "Отдел 3", "Нельзя"])
        buffer = BytesIO()
        workbook.save(buffer)

        items, skipped, conflicts = parse_mapping_workbook(buffer.getvalue())
        self.assertEqual(skipped, 2)
        self.assertEqual(conflicts, 0)
        self.assertEqual(len(items), 1)
        self.assertEqual(items[0]["kind"], "department")
        self.assertEqual(items[0]["account_code"], "25")
        self.assertEqual(items[0]["source_text"], "Отдел 1")

    def test_parallel_skips_non_fot_department(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "мэппинг"
        headers = {
            2: "Счет 1С",
            3: "Статьи затрат / Прочие доходы, расходы",
            4: "Признак",
            7: "Счет 1С",
            8: "Год",
            9: "Номенклатурные группы из 1С",
            10: "для БДР",
            13: "Счет 1С",
            14: "Год",
            15: "Признак (ТОЛЬКО ФОТ)",
            16: "Подразделение из 1С",
            17: "для БДР",
        }
        for col, value in headers.items():
            sheet.cell(row=1, column=col, value=value)
        sheet.cell(row=2, column=2, value="91.02")
        sheet.cell(row=2, column=3, value="Аренда")
        sheet.cell(row=2, column=4, value="Прочие расходы")
        sheet.cell(row=2, column=13, value="25")
        sheet.cell(row=2, column=14, value=2026)
        sheet.cell(row=2, column=15, value="Не ФОТ")
        sheet.cell(row=2, column=16, value="Отдел")
        sheet.cell(row=2, column=17, value="Нельзя")
        buffer = BytesIO()
        workbook.save(buffer)

        items, skipped, conflicts = parse_mapping_workbook(buffer.getvalue())
        self.assertEqual(skipped, 1)
        self.assertEqual(conflicts, 0)
        self.assertEqual(len(items), 1)
        self.assertEqual(items[0]["kind"], "article")

    def test_parses_unified_sheet(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Единый мэппинг"
        sheet.append(
            [
                "Счёт",
                "Год",
                "Признак ОСВ",
                "Подразделение / статья / прочее",
                "Признак БДР",
            ]
        )
        sheet.append(["90.01", 2026, "Номенклатурная группа", "Синергия", "СМР"])
        sheet.append(["91.01", "", "Статья", "Госпошлина", "Прочие доходы"])
        sheet.append(["90.02", 2026, "Подразделение", "Отдел 1", "ФОТ АУП"])
        sheet.append(["90.02", 2026, "Номенклатурная группа", "Группа", "СМР"])
        sheet.append(["91.02", "", "Статья", "Госпошлина", "Прочие расходы"])
        sheet.append(["25", 2026, "Подразделение", "Отдел 2", "Нельзя"])
        buffer = BytesIO()
        workbook.save(buffer)

        items, skipped, conflicts = parse_mapping_workbook(buffer.getvalue())
        self.assertEqual(skipped, 1)
        self.assertEqual(conflicts, 0)
        self.assertEqual(len(items), 9)

        by_key = {
            (item["kind"], item["account_code"], item["source_key"]): item
            for item in items
        }
        self.assertEqual(
            by_key[("article", "91.01", "госпошлина")]["year"], 0
        )
        self.assertEqual(
            by_key[("article", "91.01", "госпошлина")]["target"],
            "Прочие доходы",
        )
        self.assertEqual(
            by_key[("department", "90.02", "отдел 1")]["target"],
            "ФОТ АУП",
        )
        self.assertFalse(
            any(item["account_code"] == "25" for item in items)
        )
        self.assertEqual(
            by_key[("nomenclature", "90.01", "сдача имущества в аренду")][
                "target"
            ],
            "Прочее",
        )

    def test_parses_fot_sheet(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "ФОТ 26 + 90.02"
        sheet.append(
            ["Счёт", "Год", "Подразделение / признак ОСВ", "Признак БДР"]
        )
        sheet.append(["26", 2026, "Бухгалтерия", "Строители (КП)"])
        sheet.append(["90.02", 2026, "Отдел проектирования", "Проектировщики (П)"])
        sheet.append(["25", 2026, "Отдел 2", "Нельзя"])
        buffer = BytesIO()
        workbook.save(buffer)

        items, skipped, conflicts = parse_mapping_workbook(buffer.getvalue())
        self.assertEqual(len(items), 2)
        self.assertEqual(skipped, 1)
        self.assertEqual(conflicts, 0)
        self.assertTrue(all(item["kind"] == "department" for item in items))

    def test_fot_template_has_headers_and_is_empty(self):
        workbook = load_workbook(BytesIO(build_fot_mapping_template()))
        self.assertEqual(workbook.sheetnames, ["ФОТ 26 + 90.02"])
        sheet = workbook["ФОТ 26 + 90.02"]
        self.assertEqual(
            [cell.value for cell in sheet[1]],
            [
                "Счёт",
                "Год",
                "Подразделение / признак ОСВ",
                "Признак БДР",
            ],
        )
        self.assertEqual(sheet.max_row, 1)

    def test_template_has_unified_headers_and_is_empty(self):
        workbook = load_workbook(BytesIO(build_mapping_template()))
        self.assertEqual(workbook.sheetnames, ["Единый мэппинг"])
        sheet = workbook["Единый мэппинг"]
        self.assertEqual(
            [cell.value for cell in sheet[1]],
            [
                "Счёт",
                "Год",
                "Признак ОСВ",
                "Подразделение / статья / прочее",
                "Признак БДР",
            ],
        )
        self.assertEqual(sheet.max_row, 1)


if __name__ == "__main__":
    unittest.main()
