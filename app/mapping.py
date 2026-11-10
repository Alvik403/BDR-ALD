from __future__ import annotations

import re
from decimal import Decimal
from io import BytesIO

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from sqlalchemy import delete, insert, select
from sqlalchemy.orm import Session

from .models import MappingEntry, MappedRow, OsvRow, Upload

ALLOWED_MAPPING_KINDS = ("article", "nomenclature", "department")
KIND_LABELS = {
    "article": "Статьи и прочие доходы/расходы (B–D)",
    "nomenclature": "Номенклатурные группы (G–J)",
    "department": "ФОТ по подразделениям (M–Q)",
}

LABOR_ARTICLES = {
    "оплата труда",
    "оплата больничного",
    "страховые взносы",
    "взносы в фсс от нс и пз",
}
CONSTANT_TARGETS = {
    "25": "Общепроизводственные расходы",
    "26": "Общехозяйственные расходы",
}
UNNAMED = "Без названия"
ZERO = Decimal("0.00")
SPECIAL_NOMENCLATURE = (
    "Сдача имущества в аренду",
    "Реализация товаров несобственного производства",
)

UNIFIED_TEMPLATE_TITLE = "Единый мэппинг"
UNIFIED_TEMPLATE_HEADERS = (
    "Счёт",
    "Год",
    "Признак ОСВ",
    "Подразделение / статья / прочее",
    "Признак БДР",
)
OSV_KIND_LABELS = {
    "article": "Статья",
    "nomenclature": "Номенклатурная группа",
    "department": "Подразделение",
}
TEMPLATE_SHEETS = (
    ("unified", UNIFIED_TEMPLATE_TITLE, UNIFIED_TEMPLATE_HEADERS),
)
FOT_TEMPLATE_TITLE = "ФОТ 26 + 90.02"
FOT_TEMPLATE_HEADERS = (
    "Счёт",
    "Год",
    "Подразделение / признак ОСВ",
    "Признак БДР",
)


class MappingImportError(ValueError):
    """Raised when a mapping workbook cannot be imported."""


def normalize_key(value: str | None) -> str:
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ").replace("ё", "е").replace("Ё", "Е")
    text = re.sub(r"[\"'«»„“”]", "", text)
    text = text.replace("–", "-").replace("—", "-")
    text = re.sub(r"\s+", " ", text).strip().casefold()
    text = re.sub(r"[\s.,;:]+$", "", text)
    return text


def extract_year(period: str | None) -> int:
    if not period:
        return 0
    match = re.search(r"(\d{4})", str(period))
    return int(match.group(1)) if match else 0


def load_maps(
    db: Session,
) -> dict[str, dict[tuple[str, int, str], str]]:
    maps: dict[str, dict[tuple[str, int, str], str]] = {
        kind: {} for kind in ALLOWED_MAPPING_KINDS
    }
    for entry in db.scalars(select(MappingEntry)).all():
        if entry.kind not in maps:
            continue
        target = (entry.target or "").strip()
        if target:
            maps[entry.kind][
                (
                    entry.account_code or "",
                    int(entry.year or 0),
                    entry.source_key,
                )
            ] = target
    return maps


def leaf_dimensions(
    rows: list[OsvRow], account_code: str
) -> list[dict]:
    """Flatten leaf ОСВ rows into department / group / article dimensions."""
    parent_ids = {row.parent_id for row in rows if row.parent_id is not None}
    by_id = {row.id: row for row in rows}
    result: list[dict] = []

    for row in rows:
        if row.id in parent_ids:
            continue
        if row.node_type in ("account", "division", "total"):
            continue

        ancestors: list[OsvRow] = []
        parent = by_id.get(row.parent_id)
        while parent is not None:
            if parent.node_type != "account":
                ancestors.append(parent)
            parent = by_id.get(parent.parent_id)
        ancestors.reverse()

        department = ancestors[0].name if ancestors else ""
        immediate = (
            ancestors[-1]
            if row.node_type == "month" and ancestors
            else row
        )
        immediate_name = immediate.name if immediate else ""

        if account_code == "90.01":
            group_ancestor = (
                ancestors[1]
                if len(ancestors) >= 2
                else (immediate if row.node_type == "month" else row)
            )
            group_name = group_ancestor.name if group_ancestor else ""
            article = ""
        elif account_code == "90.02":
            if len(ancestors) >= 2:
                group_name = ancestors[1].name
            else:
                group_name = row.name if row.node_type != "month" else ""
            if row.node_type == "month":
                article = immediate_name
            elif row.level >= 3:
                article = row.name
            else:
                article = ""
        else:
            group_name = ""
            article = immediate_name

        result.append(
            {
                "row": row,
                "department": department or UNNAMED,
                "group_name": group_name or "",
                "article": article or "",
            }
        )

    return result


def resolve_mapping(
    account_code: str,
    department: str,
    group_name: str,
    article: str,
    maps: dict[str, dict[tuple[str, int, str], str]],
    target_account: str = "",
    target_year: int = 0,
) -> tuple[str | None, str | None, str | None]:
    accounts: list[str] = []
    candidates = [target_account or ""]
    if target_account and "." in target_account:
        candidates.append(target_account.split(".", 1)[0])
    candidates.append("")
    for candidate in candidates:
        if candidate not in accounts:
            accounts.append(candidate)

    years: list[int] = []
    for candidate in (int(target_year or 0), 0):
        if candidate not in years:
            years.append(candidate)

    def lookup(kind: str, key: str) -> str | None:
        normalized = normalize_key(key)
        for acc in accounts:
            for year in years:
                target = maps.get(kind, {}).get((acc, year, normalized))
                if target:
                    target = target.strip()
                    if target:
                        return target
        return None

    if account_code == "90.01":
        key = group_name or UNNAMED
        return lookup("nomenclature", key), "nomenclature", key

    if account_code == "90.02":
        if normalize_key(article) in LABOR_ARTICLES:
            key = department or UNNAMED
            return lookup("department", key), "department", key
        key = group_name or UNNAMED
        return lookup("nomenclature", key), "nomenclature", key

    if account_code in ("25", "26"):
        if normalize_key(article) in LABOR_ARTICLES:
            key = department or UNNAMED
            return lookup("department", key), "department", key
        return CONSTANT_TARGETS[account_code], "constant", None

    if account_code in ("91.01", "91.02"):
        key = article or UNNAMED
        return lookup("article", key), "article", key

    return None, None, None


def build_mapped_rows(
    upload: Upload,
    rows: list[OsvRow],
    maps: dict[str, dict[tuple[str, int, str], str]],
) -> list[dict]:
    payload: list[dict] = []
    upload_year = extract_year(upload.period)
    for dims in leaf_dimensions(rows, upload.account_code):
        row = dims["row"]
        target, kind, key = resolve_mapping(
            upload.account_code,
            dims["department"],
            dims["group_name"],
            dims["article"],
            maps,
            upload.account_code,
            upload_year,
        )
        payload.append(
            {
                "upload_id": upload.id,
                "osv_row_id": row.id,
                "account_code": upload.account_code,
                "position": row.position,
                "department": dims["department"],
                "group_name": dims["group_name"],
                "article": dims["article"],
                "month": row.month,
                "debit": row.turnover_debit or ZERO,
                "credit": row.turnover_credit or ZERO,
                "target": target,
                "mapping_kind": kind,
                "mapping_key": key,
                "status": "mapped" if target else "unmapped",
            }
        )
    return payload


def rebuild_mapped_rows(db: Session, upload_id: str | None = None) -> int:
    """Rebuild mapped rows for one upload or for every upload."""
    maps = load_maps(db)
    query = select(Upload)
    if upload_id is not None:
        query = query.where(Upload.id == upload_id)
    uploads = db.scalars(query).all()

    total = 0
    for upload in uploads:
        rows = db.scalars(
            select(OsvRow)
            .where(OsvRow.upload_id == upload.id)
            .order_by(OsvRow.position)
        ).all()
        payload = build_mapped_rows(upload, rows, maps)
        db.execute(delete(MappedRow).where(MappedRow.upload_id == upload.id))
        if payload:
            db.execute(insert(MappedRow), payload)
        total += len(payload)
    db.commit()
    return total


def kinds_for_account(code: str | None) -> list[str]:
    if code == "90.01":
        return ["nomenclature"]
    if code == "90.02":
        return ["nomenclature", "department"]
    if code in ("25", "26"):
        return ["department"]
    if code in ("91.01", "91.02"):
        return ["article"]
    return list(ALLOWED_MAPPING_KINDS)


def _style_unified_sheet(sheet, rows_count: int) -> None:
    header_font = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="EEF3FF")
    for cell in sheet[1]:
        cell.font = header_font
        cell.fill = header_fill
    sheet.freeze_panes = "A2"
    sheet.column_dimensions["A"].width = 14
    sheet.column_dimensions["A"].number_format = "@"
    sheet.column_dimensions["B"].width = 10
    sheet.column_dimensions["C"].width = 24
    sheet.column_dimensions["D"].width = 58
    sheet.column_dimensions["E"].width = 30
    if rows_count:
        sheet.auto_filter.ref = f"A1:E{rows_count + 1}"
    kind_validation = DataValidation(
        type="list",
        formula1='"Статья,Номенклатурная группа,Подразделение"',
        allow_blank=True,
    )
    account_validation = DataValidation(
        type="list",
        formula1='"26,90.01,90.02,91.01,91.02"',
        allow_blank=True,
    )
    sheet.add_data_validation(kind_validation)
    sheet.add_data_validation(account_validation)
    kind_validation.add("C2:C5000")
    account_validation.add("A2:A5000")


def build_mapping_template() -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = UNIFIED_TEMPLATE_TITLE
    sheet.append(list(UNIFIED_TEMPLATE_HEADERS))
    _style_unified_sheet(sheet, 0)
    buffer = BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def build_mapping_export(entries: list[MappingEntry]) -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = UNIFIED_TEMPLATE_TITLE
    rows = []
    kind_order = {"article": 0, "nomenclature": 1, "department": 2}
    for entry in entries:
        kind = entry.kind
        if kind not in OSV_KIND_LABELS:
            continue
        if kind == "department" and entry.account_code == "25":
            continue
        target = (entry.target or "").strip()
        if not target:
            continue
        year = int(entry.year or 0)
        rows.append(
            (
                entry.account_code or "",
                year if year else "",
                OSV_KIND_LABELS[kind],
                entry.source_text,
                target,
                kind_order.get(kind, 9),
            )
        )
    rows.sort(
        key=lambda row: (
            row[0],
            row[5],
            (row[3] or "").casefold(),
        )
    )
    sheet.append(list(UNIFIED_TEMPLATE_HEADERS))
    for row in rows:
        sheet.append(list(row[:5]))
    _style_unified_sheet(sheet, len(rows))
    buffer = BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()
def _style_fot_sheet(sheet, rows_count: int) -> None:
    header_font = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="EEF3FF")
    for cell in sheet[1]:
        cell.font = header_font
        cell.fill = header_fill
    sheet.freeze_panes = "A2"
    sheet.column_dimensions["A"].width = 14
    sheet.column_dimensions["A"].number_format = "@"
    sheet.column_dimensions["B"].width = 10
    sheet.column_dimensions["C"].width = 58
    sheet.column_dimensions["D"].width = 30
    if rows_count:
        sheet.auto_filter.ref = f"A1:D{rows_count + 1}"
    account_validation = DataValidation(
        type="list",
        formula1='"26,90.02"',
        allow_blank=True,
    )
    sheet.add_data_validation(account_validation)
    account_validation.add("A2:A5000")


def build_fot_mapping_template() -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = FOT_TEMPLATE_TITLE
    sheet.append(list(FOT_TEMPLATE_HEADERS))
    _style_fot_sheet(sheet, 0)
    buffer = BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def build_fot_mapping_export(rows: list[dict]) -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = FOT_TEMPLATE_TITLE
    sheet.append(list(FOT_TEMPLATE_HEADERS))
    ordered = sorted(
        rows,
        key=lambda row: (
            str(row.get("account") or ""),
            str(row.get("department") or "").casefold(),
        ),
    )
    for row in ordered:
        sheet.append(
            [
                row.get("account") or "",
                row.get("year") or "",
                row.get("department") or "",
                row.get("target") or "",
            ]
        )
    _style_fot_sheet(sheet, len(ordered))
    buffer = BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def _sheet_kind(title: str) -> str | None:
    text = normalize_key(title)
    if "стат" in text or "article" in text:
        return "article"
    if "номен" in text or "nomenclature" in text:
        return "nomenclature"
    if "подраз" in text or "department" in text:
        return "department"
    return None


_SHEET_COLUMNS = {
    "article": {
        "account": ("счет 1с",),
        "key": ("статьи затрат", "прочие доходы"),
        "target": ("признак",),
    },
    "nomenclature": {
        "account": ("счет 1с",),
        "year": ("год",),
        "key": ("номенклатурные группы", "номенклатурная группа"),
        "target": ("для бдр", "бдр"),
    },
    "department": {
        "account": ("счет 1с",),
        "year": ("год",),
        "flag": ("признак",),
        "key": ("подразделение",),
        "target": ("для бдр", "бдр"),
    },
}
_FALLBACK_COLUMNS = {
    "article": {"account": 1, "key": 2, "target": 3},
    "nomenclature": {"account": 1, "year": 2, "key": 3, "target": 4},
    "department": {
        "account": 1,
        "year": 2,
        "flag": 3,
        "key": 4,
        "target": 5,
    },
}
_HEADER_KEYS = {
    "счет 1с",
    "счет",
    "счёт",
    "год",
    "признак осв",
    "подразделение / статья / прочее",
    "подразделение / признак осв",
    "признак бдр",
    "статьи затрат / прочие доходы, расходы",
    "номенклатурные группы из 1с",
    "подразделение из 1с",
}


def _sheet_columns(ws, kind: str) -> tuple[dict[str, int], int]:
    config = _SHEET_COLUMNS[kind]
    for row in range(1, min(ws.max_row or 0, 10) + 1):
        cols: dict[str, int] = {}
        for col in range(1, min(ws.max_column or 0, 15) + 1):
            text = normalize_key(ws.cell(row=row, column=col).value)
            if not text:
                continue
            for field, needles in config.items():
                if field in cols:
                    continue
                if any(needle in text for needle in needles):
                    cols[field] = col
        if "key" in cols and "target" in cols:
            return cols, row + 1
    return dict(_FALLBACK_COLUMNS[kind]), 1


def _detect_kind(ws, title: str) -> str | None:
    kind = _sheet_kind(title)
    if kind:
        return kind
    for row in range(1, min(ws.max_row or 0, 10) + 1):
        for col in range(1, min(ws.max_column or 0, 15) + 1):
            text = normalize_key(ws.cell(row=row, column=col).value)
            if not text:
                continue
            if "номенклатурные группы" in text:
                return "nomenclature"
            if "подразделение" in text:
                return "department"
            if "статьи затрат" in text or "прочие доходы" in text:
                return "article"
    return None


def _cell_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _is_placeholder(value) -> bool:
    return normalize_key(value) in ("", "<...>", "...", "…")


def _is_fot_flag(value: str) -> bool:
    text = normalize_key(value)
    if not text:
        return True
    if "нефот" in text or "безфот" in text or "без фот" in text:
        return False
    if re.search(r"(^|\s)не\s*-?\s*фот", text):
        return False
    return "фот" in text


_UNIFIED_COLUMN_ALIASES = {
    "account": ("счет 1с", "счет", "счёт 1с", "счёт"),
    "year": ("год",),
    "kind": ("признак осв", "тип осв", "вид осв"),
    "source": (
        "подразделение / статья / прочее",
        "подразделение/статья/прочее",
        "значение из осв",
        "подразделение из 1с",
        "номенклатурные группы из 1с",
        "статьи затрат / прочие доходы, расходы",
    ),
    "target": ("признак бдр", "для бдр", "цель бдр"),
}


def _kind_from_osv(value) -> str | None:
    text = normalize_key(value)
    if not text:
        return None
    if (
        "номенклатур" in text
        or "nomenclature" in text
        or "g-j" in text
        or "g j" in text
    ):
        return "nomenclature"
    if (
        "подраздел" in text
        or "department" in text
        or "m-q" in text
        or "m q" in text
        or "фот" in text
    ):
        return "department"
    if (
        "стать" in text
        or "article" in text
        or "b-d" in text
        or "b d" in text
    ):
        return "article"
    return None


def _unified_header_row(ws) -> tuple[int, dict[str, int]] | None:
    for row in range(1, min(ws.max_row or 0, 15) + 1):
        cols: dict[str, int] = {}
        for col in range(1, min(ws.max_column or 0, 20) + 1):
            text = normalize_key(ws.cell(row=row, column=col).value)
            if not text:
                continue
            for field, aliases in _UNIFIED_COLUMN_ALIASES.items():
                if field in cols:
                    continue
                if any(alias in text for alias in aliases):
                    cols[field] = col
        if all(field in cols for field in ("account", "kind", "source", "target")):
            kind_text = normalize_key(
                ws.cell(row=row, column=cols["kind"]).value
            )
            if "осв" in kind_text:
                return row, cols
    return None


def _parse_unified_sheet(ws) -> tuple[list[dict], int, int] | None:
    header = _unified_header_row(ws)
    if header is None:
        return None
    header_row, cols = header

    allowed_accounts = {
        "article": {"91.01", "91.02"},
        "nomenclature": {"90.01", "90.02"},
        "department": {"26", "90.02"},
    }
    items_by_key: dict[tuple[str, str, int, str], dict] = {}
    skipped = 0
    conflicts = 0

    for row in range(header_row + 1, (ws.max_row or 0) + 1):
        account_text = _cell_text(
            ws.cell(row=row, column=cols["account"]).value
        )
        year_text = (
            _cell_text(ws.cell(row=row, column=cols["year"]).value)
            if "year" in cols
            else ""
        )
        kind = _kind_from_osv(
            ws.cell(row=row, column=cols["kind"]).value
        )
        source_text = _cell_text(
            ws.cell(row=row, column=cols["source"]).value
        )
        target_text = _cell_text(
            ws.cell(row=row, column=cols["target"]).value
        )
        if not account_text and not source_text and not target_text:
            continue
        if not kind or account_text not in allowed_accounts.get(kind, set()):
            skipped += 1
            continue
        if not source_text or not target_text:
            skipped += 1
            continue
        source_key = normalize_key(source_text)
        if not source_key or source_key in _HEADER_KEYS:
            continue
        year = 0
        if kind != "article" and year_text:
            digits = "".join(char for char in year_text if char.isdigit())
            year = int(digits) if digits else 0
        payload: dict = {"account": account_text}
        if year:
            payload["year"] = year
        item_key = (kind, account_text, year, source_key)
        existing = items_by_key.get(item_key)
        if existing is not None and existing["target"] != target_text:
            conflicts += 1
        items_by_key[item_key] = {
            "kind": kind,
            "account_code": account_text,
            "year": year,
            "source_text": source_text,
            "source_key": source_key,
            "target": target_text,
            "payload": payload,
        }

    if not items_by_key:
        raise MappingImportError(
            "В файле нет заполненных строк единого мэппинга. "
            "Заполните колонки «Признак ОСВ», «Подразделение / статья / прочее» "
            "и «Признак БДР»."
        )
    _add_special_overrides(items_by_key)
    return list(items_by_key.values()), skipped, conflicts


_FOT_COLUMN_ALIASES = {
    "account": ("счет 1с", "счет", "счёт 1с", "счёт"),
    "year": ("год",),
    "source": (
        "подразделение / признак осв",
        "подразделение/признак осв",
    ),
    "target": ("признак бдр", "для бдр", "цель бдр"),
}


def _fot_header_row(ws) -> tuple[int, dict[str, int]] | None:
    for row in range(1, min(ws.max_row or 0, 15) + 1):
        cols: dict[str, int] = {}
        for col in range(1, min(ws.max_column or 0, 12) + 1):
            text = normalize_key(ws.cell(row=row, column=col).value)
            if not text:
                continue
            for field, aliases in _FOT_COLUMN_ALIASES.items():
                if field in cols:
                    continue
                if any(alias in text for alias in aliases):
                    cols[field] = col
        if all(field in cols for field in ("account", "source", "target")):
            source_text = normalize_key(
                ws.cell(row=row, column=cols["source"]).value
            )
            if "подраздел" in source_text:
                return row, cols
    return None


def _parse_fot_sheet(ws) -> tuple[list[dict], int, int] | None:
    header = _fot_header_row(ws)
    if header is None:
        return None
    header_row, cols = header
    items_by_key: dict[tuple[str, str, int, str], dict] = {}
    skipped = 0
    conflicts = 0
    for row in range(header_row + 1, (ws.max_row or 0) + 1):
        account_text = _cell_text(
            ws.cell(row=row, column=cols["account"]).value
        )
        year_text = (
            _cell_text(ws.cell(row=row, column=cols["year"]).value)
            if "year" in cols
            else ""
        )
        department = _cell_text(
            ws.cell(row=row, column=cols["source"]).value
        )
        target_text = _cell_text(
            ws.cell(row=row, column=cols["target"]).value
        )
        if not account_text and not department and not target_text:
            continue
        if account_text not in ("26", "90.02") or not department or not target_text:
            skipped += 1
            continue
        source_key = normalize_key(department)
        if not source_key or source_key in _HEADER_KEYS:
            continue
        year = 0
        if year_text:
            digits = "".join(char for char in year_text if char.isdigit())
            year = int(digits) if digits else 0
        payload: dict = {"account": account_text}
        if year:
            payload["year"] = year
        item_key = ("department", account_text, year, source_key)
        existing = items_by_key.get(item_key)
        if existing is not None and existing["target"] != target_text:
            conflicts += 1
        items_by_key[item_key] = {
            "kind": "department",
            "account_code": account_text,
            "year": year,
            "source_text": department,
            "source_key": source_key,
            "target": target_text,
            "payload": payload,
        }
    if not items_by_key:
        raise MappingImportError(
            "В файле нет заполненных строк ФОТ. Заполните колонки "
            "«Подразделение / признак ОСВ» и «Признак БДР»."
        )
    return list(items_by_key.values()), skipped, conflicts


def _add_special_overrides(items_by_key: dict[tuple, dict]) -> None:
    if not any(
        item["kind"] == "nomenclature"
        for item in items_by_key.values()
    ):
        return
    max_year = max(
        (item["year"] for item in items_by_key.values() if item.get("year")),
        default=0,
    )
    for account_code in ("90.01", "90.02"):
        for source_text in SPECIAL_NOMENCLATURE:
            source_key = normalize_key(source_text)
            payload = {"account": account_code}
            if max_year:
                payload["year"] = max_year
            items_by_key[
                ("nomenclature", account_code, max_year, source_key)
            ] = {
                "kind": "nomenclature",
                "account_code": account_code,
                "year": max_year,
                "source_text": source_text,
                "source_key": source_key,
                "target": "Прочее",
                "payload": payload,
            }


def _parallel_header_row(ws) -> int | None:
    for row in range(1, min(ws.max_row or 0, 20) + 1):
        b = normalize_key(ws.cell(row=row, column=2).value)
        c = normalize_key(ws.cell(row=row, column=3).value)
        d = normalize_key(ws.cell(row=row, column=4).value)
        g = normalize_key(ws.cell(row=row, column=7).value)
        i = normalize_key(ws.cell(row=row, column=9).value)
        m = normalize_key(ws.cell(row=row, column=13).value)
        p = normalize_key(ws.cell(row=row, column=16).value)
        q = normalize_key(ws.cell(row=row, column=17).value)
        if (
            "счет 1с" in b
            and ("статьи затрат" in c or "прочие доходы" in c)
            and "признак" in d
            and "счет 1с" in g
            and "номенклатурные группы" in i
            and "счет 1с" in m
            and "подраздел" in p
            and "бдр" in q
        ):
            return row
    return None


def _parse_parallel_sheet(
    ws,
) -> tuple[list[dict], int, int] | None:
    """Parse one sheet with B–D / G–J / M–Q ranges side by side."""
    header_row = _parallel_header_row(ws)
    if header_row is None:
        return None

    start_row = header_row + 1
    max_year = 0
    for row in range(start_row, (ws.max_row or 0) + 1):
        for col in (8, 14):
            text = _cell_text(ws.cell(row=row, column=col).value)
            digits = "".join(char for char in text if char.isdigit())
            if digits:
                max_year = max(max_year, int(digits))

    items_by_key: dict[tuple[str, str, int, str], dict] = {}
    skipped = 0
    conflicts = 0

    def add(
        kind: str,
        account_code: str,
        year: int,
        key_value,
        target_value,
        flag: str = "",
    ) -> None:
        nonlocal skipped, conflicts
        if _is_placeholder(key_value):
            return
        if kind == "department":
            if account_code not in ("25", "26", "90.02"):
                skipped += 1
                return
            if flag and not _is_fot_flag(flag):
                skipped += 1
                return
        target_text = _cell_text(target_value)
        if not target_text or _is_placeholder(target_text):
            skipped += 1
            return
        key_text = _cell_text(key_value)
        source_key = normalize_key(key_text)
        if not source_key or source_key in _HEADER_KEYS:
            return
        payload: dict = {"account": account_code}
        if year:
            payload["year"] = year
        if flag:
            payload["flag"] = flag
        item_key = (kind, account_code, year, source_key)
        existing = items_by_key.get(item_key)
        if existing is not None and existing["target"] != target_text:
            conflicts += 1
        items_by_key[item_key] = {
            "kind": kind,
            "account_code": account_code,
            "year": year,
            "source_text": key_text,
            "source_key": source_key,
            "target": target_text,
            "payload": payload,
        }

    for row in range(start_row, (ws.max_row or 0) + 1):
        account = _cell_text(ws.cell(row=row, column=2).value)
        if account in ("91.01", "91.02"):
            add(
                "article",
                account,
                0,
                ws.cell(row=row, column=3).value,
                ws.cell(row=row, column=4).value,
            )

        account = _cell_text(ws.cell(row=row, column=7).value)
        year_text = _cell_text(ws.cell(row=row, column=8).value)
        digits = "".join(char for char in year_text if char.isdigit())
        year = int(digits) if digits else 0
        if account in ("90.01", "90.02") and year == max_year:
            add(
                "nomenclature",
                account,
                year,
                ws.cell(row=row, column=9).value,
                ws.cell(row=row, column=10).value,
            )

        account = _cell_text(ws.cell(row=row, column=13).value)
        year_text = _cell_text(ws.cell(row=row, column=14).value)
        digits = "".join(char for char in year_text if char.isdigit())
        year = int(digits) if digits else 0
        if account in ("25", "26", "90.02") and year == max_year:
            add(
                "department",
                account,
                year,
                ws.cell(row=row, column=16).value,
                ws.cell(row=row, column=17).value,
                _cell_text(ws.cell(row=row, column=15).value),
            )

    if not items_by_key:
        raise MappingImportError(
            "В файле не найдено заполненных строк B–D / G–J / M–Q."
        )
    _add_special_overrides(items_by_key)
    return list(items_by_key.values()), skipped, conflicts


def parse_mapping_workbook(
    data: bytes,
) -> tuple[list[dict], int, int]:
    try:
        workbook = load_workbook(BytesIO(data), data_only=True)
    except Exception as exc:  # noqa: BLE001 - openpyxl raises various errors
        raise MappingImportError(
            "Не удалось открыть файл. Нужен файл .xlsx или .xlsm."
        ) from exc

    sheets = list(workbook.worksheets)
    if not sheets:
        raise MappingImportError("В файле нет листов.")

    for sheet in sheets:
        unified = _parse_unified_sheet(sheet)
        if unified is not None:
            return unified

    for sheet in sheets:
        fot = _parse_fot_sheet(sheet)
        if fot is not None:
            return fot

    for sheet in sheets:
        parallel = _parse_parallel_sheet(sheet)
        if parallel is not None:
            return parallel

    recognized = []
    for sheet in sheets:
        kind = _detect_kind(sheet, sheet.title)
        if kind:
            recognized.append((sheet, kind))
    if not recognized:
        if len(sheets) >= 3:
            recognized = [
                (sheets[0], "article"),
                (sheets[1], "nomenclature"),
                (sheets[2], "department"),
            ]
        else:
            raise MappingImportError(
                "Не удалось определить тип листа. Используйте шаблон: "
                "листы B–D, G–J, M–Q."
            )

    items_by_key: dict[tuple[str, str, int, str], dict] = {}
    skipped = 0
    conflicts = 0
    for sheet, kind in recognized:
        cols, start_row = _sheet_columns(sheet, kind)
        for row in range(start_row, (sheet.max_row or 0) + 1):
            key_text = _cell_text(
                sheet.cell(row=row, column=cols["key"]).value
            )
            target_text = _cell_text(
                sheet.cell(row=row, column=cols["target"]).value
            )
            if not key_text:
                continue
            if not target_text:
                skipped += 1
                continue
            source_key = normalize_key(key_text)
            if not source_key:
                skipped += 1
                continue
            if source_key in _HEADER_KEYS:
                continue

            account_text = (
                _cell_text(sheet.cell(row=row, column=cols["account"]).value)
                if "account" in cols
                else ""
            )
            year_text = (
                _cell_text(sheet.cell(row=row, column=cols["year"]).value)
                if "year" in cols
                else ""
            )
            flag_text = (
                _cell_text(sheet.cell(row=row, column=cols["flag"]).value)
                if "flag" in cols
                else ""
            )
            if kind == "department":
                if account_text not in ("25", "26", "90.02"):
                    skipped += 1
                    continue
                if not _is_fot_flag(flag_text):
                    skipped += 1
                    continue
            year = 0
            if year_text:
                digits = "".join(char for char in year_text if char.isdigit())
                year = int(digits) if digits else 0

            payload = {}
            if account_text:
                payload["account"] = account_text
            if year:
                payload["year"] = year
            if flag_text:
                payload["flag"] = flag_text

            item_key = (kind, account_text, year, source_key)
            existing = items_by_key.get(item_key)
            if existing is not None and existing["target"] != target_text:
                conflicts += 1
            items_by_key[item_key] = {
                "kind": kind,
                "account_code": account_text,
                "year": year,
                "source_text": key_text,
                "source_key": source_key,
                "target": target_text,
                "payload": payload or None,
            }

    items = list(items_by_key.values())
    if not items:
        raise MappingImportError(
            "В файле нет заполненных строк: заполните колонки "
            "«для БДР» и «Признак»."
        )
    _add_special_overrides(items_by_key)
    return list(items_by_key.values()), skipped, conflicts


def apply_mapping_entries(
    db: Session, items: list[dict]
) -> tuple[int, int]:
    existing = {
        (
            entry.kind,
            entry.account_code or "",
            int(entry.year or 0),
            entry.source_key,
        ): entry
        for entry in db.scalars(select(MappingEntry)).all()
    }
    added = updated = 0
    for item in items:
        account_code = str(item.get("account_code") or "").strip()
        year = int(item.get("year") or 0)
        key = (
            item["kind"],
            account_code,
            year,
            item["source_key"],
        )
        entry = existing.get(key)
        if entry is None:
            entry = MappingEntry(
                kind=item["kind"],
                account_code=account_code,
                year=year,
                source_key=item["source_key"],
                source_text=item["source_text"],
                target=item["target"],
                payload=item.get("payload"),
            )
            db.add(entry)
            existing[key] = entry
            added += 1
        else:
            entry.source_text = item["source_text"]
            entry.target = item["target"]
            entry.payload = item.get("payload")
            updated += 1
    db.commit()
    return added, updated
