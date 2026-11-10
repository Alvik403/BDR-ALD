from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any

import openpyxl

SUPPORTED_ACCOUNTS = ("25", "26", "90.01", "90.02", "91.01", "91.02")
ZERO = Decimal("0.00")

_ACCOUNT_RE = re.compile(r"по\s+сч[её]ту\s*№?\s*(\d{2}(?:\.\d{1,2})?)", re.IGNORECASE)
_FILE_ACCOUNT_RE = re.compile(r"^\s*(\d{2}(?:\.\d{1,2})?)\s*[_\-.]")
_GROUP_LABELS = {
    "opening": "сальдо на начало",
    "turnover": "обороты за период",
    "closing": "сальдо на конец",
}
_NODE_LEVELS = {"account": 0, "division": 1, "item": 2, "total": 2, "month": 3}


class ParseError(ValueError):
    """Raised when an uploaded file is not a recognizable ОСВ report."""


@dataclass
class ParsedRow:
    position: int
    level: int
    node_type: str
    code: str | None
    name: str
    month: str | None
    amounts: tuple[Decimal, Decimal, Decimal, Decimal, Decimal, Decimal]
    parent_index: int | None = None


@dataclass
class ParsedReport:
    account_code: str
    company: str | None
    period: str | None
    source: str | None
    sheet_name: str
    rows: list[ParsedRow] = field(default_factory=list)
    totals: tuple[Decimal, Decimal, Decimal, Decimal, Decimal, Decimal] = field(
        default=(ZERO, ZERO, ZERO, ZERO, ZERO, ZERO)
    )


def _norm(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ").strip().lower()
    return re.sub(r"\s+", " ", text)


def _text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _decimal(value: Any) -> Decimal | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return Decimal(value).quantize(Decimal("0.01"))
    if isinstance(value, float):
        return Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    if isinstance(value, Decimal):
        return value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    text = str(value).strip().replace("\xa0", "").replace(" ", "").replace(",", ".")
    if text in ("", "-", "—", "x", "х"):
        return None
    try:
        return Decimal(text).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError):
        return None


def _indent(ws, row: int) -> int:
    cell = ws.cell(row=row, column=1)
    indent = cell.alignment.indent if cell.alignment else None
    if not indent:
        dim = ws.row_dimensions.get(row)
        if dim is not None and dim.outlineLevel:
            indent = dim.outlineLevel
    try:
        return int(round(float(indent or 0)))
    except (TypeError, ValueError):
        return 0


def _pick_sheet(wb):
    for ws in wb.worksheets:
        for row in range(1, min(ws.max_row, 25) + 1):
            if "оборотно-сальдовая" in _norm(ws.cell(row=row, column=1).value):
                return ws
    return wb.worksheets[0]


def _find_report_meta(ws) -> tuple[str | None, int | None, str | None]:
    """Return (title text, title row, account code from the title)."""
    for row in range(1, min(ws.max_row, 60) + 1):
        for col in range(1, min(ws.max_column, 5) + 1):
            text = _text(ws.cell(row=row, column=col).value)
            if not text:
                continue
            norm = text.lower()
            if "оборотно-сальдовая" in norm:
                match = _ACCOUNT_RE.search(text)
                return text, row, match.group(1) if match else None
    for row in range(1, min(ws.max_row, 60) + 1):
        text = _text(ws.cell(row=row, column=1).value)
        match = _ACCOUNT_RE.search(text) if text else None
        if match:
            return text, row, match.group(1)
    return None, None, None


def _find_headers(ws) -> tuple[dict[str, tuple[int, int]], int, int]:
    """Locate the ОСВ header block.

    Returns (column map, data start row, last header row).
    Column map: group -> (debit column, credit column).
    """
    best_row = None
    best_found: dict[str, int] = {}
    for row in range(1, min(ws.max_row, 80) + 1):
        found: dict[str, int] = {}
        for col in range(1, min(ws.max_column, 30) + 1):
            value = _norm(ws.cell(row=row, column=col).value)
            if not value:
                continue
            for key, label in _GROUP_LABELS.items():
                if label in value and key not in found:
                    found[key] = col
        if len(found) > len(best_found):
            best_row, best_found = row, found

    if best_row is None or len(best_found) < 2:
        raise ParseError(
            "Не удалось найти шапку ОСВ. Ожидались колонки «Сальдо на начало», "
            "«Обороты за период» и «Сальдо на конец»."
        )

    merged = list(ws.merged_cells.ranges)
    column_map: dict[str, tuple[int, int]] = {}
    sub_rows: list[int] = []

    for key, group_col in best_found.items():
        max_col = group_col + 2
        for rng in merged:
            if rng.min_row == best_row and rng.min_col == group_col:
                max_col = max(max_col, rng.max_col)
                break

        debit = credit = None
        used_sub_row = None
        for sub_row in range(best_row + 1, min(best_row + 3, ws.max_row) + 1):
            for col in range(group_col, max_col + 1):
                value = _norm(ws.cell(row=sub_row, column=col).value)
                if value == "дебет" and debit is None:
                    debit = col
                elif value == "кредит" and credit is None:
                    credit = col
            if debit is not None and credit is not None:
                used_sub_row = sub_row
                break
        if debit is None or credit is None:
            raise ParseError(
                "Не удалось разобрать колонки «Дебет/Кредит» в шапке ОСВ."
            )
        column_map[key] = (debit, credit)
        sub_rows.append(used_sub_row)

    if len(column_map) < 3:
        raise ParseError(
            "В шапке ОСВ не хватает одной из групп колонок "
            "(сальдо на начало / обороты / сальдо на конец)."
        )

    last_header_row = max(sub_rows)
    amount_cols = [col for pair in column_map.values() for col in pair]
    data_start = None
    for row in range(last_header_row + 1, ws.max_row + 1):
        if any(ws.cell(row=row, column=col).value is not None for col in amount_cols):
            data_start = row
            break
    if data_start is None:
        raise ParseError("В файле не найдено ни одной строки данных ОСВ.")

    return column_map, data_start, last_header_row


def _detect_account(code: str | None, filename: str | None) -> str:
    candidates = []
    if code:
        candidates.append(code.strip())
    if filename:
        match = _FILE_ACCOUNT_RE.search(filename)
        if match:
            candidates.append(match.group(1))
        else:
            match = re.search(r"(\d{2}(?:\.\d{2})?)", filename)
            if match:
                candidates.append(match.group(1))
    for candidate in candidates:
        normalized = candidate.replace(",", ".")
        if normalized in SUPPORTED_ACCOUNTS:
            return normalized
    if candidates:
        raise ParseError(
            f"Счёт {candidates[0]} не поддерживается. Доступны: "
            + ", ".join(SUPPORTED_ACCOUNTS)
            + "."
        )
    raise ParseError(
        "Не удалось определить счёт. Укажите его в заголовке отчёта "
        "(«… по счету 25 …») или в имени файла (например, 25_ОСВ.xlsx)."
    )


def parse_osv(data: bytes, filename: str | None = None) -> ParsedReport:
    try:
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    except Exception as exc:  # noqa: BLE001 - openpyxl raises various types
        raise ParseError(
            "Не удалось открыть файл Excel. Нужен файл формата .xlsx или .xlsm."
        ) from exc

    ws = _pick_sheet(wb)
    if ws.max_row is None or ws.max_row < 2:
        raise ParseError("Лист Excel пуст.")

    title, title_row, title_account = _find_report_meta(ws)
    account_code = _detect_account(title_account, filename)

    company = None
    if title_row:
        for row in range(1, title_row):
            text = _text(ws.cell(row=row, column=1).value)
            if text:
                company = text
                break

    source = None
    for row in range(1, min(ws.max_row, 30) + 1):
        for col in range(1, min(ws.max_column, 5) + 1):
            text = _text(ws.cell(row=row, column=col).value)
            if "выводимые данные" in text.lower():
                source = text
                break
        if source:
            break

    period = None
    if title:
        match = re.search(r"\bза\s+(.+)$", title, re.IGNORECASE)
        if match:
            period = re.sub(r"\s+", " ", match.group(1)).strip()

    column_map, data_start, _ = _find_headers(ws)
    amount_cols = [col for pair in column_map.values() for col in pair]

    raw_rows: list[dict[str, Any]] = []
    for row in range(data_start, ws.max_row + 1):
        cell_value = ws.cell(row=row, column=1).value
        amounts_by_group = {}
        any_amount = False
        for key, (debit_col, credit_col) in column_map.items():
            debit = _decimal(ws.cell(row=row, column=debit_col).value)
            credit = _decimal(ws.cell(row=row, column=credit_col).value)
            if debit is not None or credit is not None:
                any_amount = True
            amounts_by_group[key] = (debit or ZERO, credit or ZERO)
        if not any_amount:
            continue

        text = _text(cell_value)
        if isinstance(cell_value, (int, float, Decimal)) and not isinstance(
            cell_value, bool
        ):
            node_type = "division"
        elif re.match(r"^обороты\s+за\b", text, re.IGNORECASE):
            node_type = "month"
        else:
            node_type = "item"

        raw_rows.append(
            {
                "row": row,
                "text": text,
                "raw": cell_value,
                "node_type": node_type,
                "indent": _indent(ws, row),
                "amounts": (
                    amounts_by_group["opening"][0],
                    amounts_by_group["opening"][1],
                    amounts_by_group["turnover"][0],
                    amounts_by_group["turnover"][1],
                    amounts_by_group["closing"][0],
                    amounts_by_group["closing"][1],
                ),
            }
        )

    if not raw_rows:
        raise ParseError("В файле не найдено строк с оборотами.")

    # The first data row is the account total; account code rows are recognized too.
    for index, raw in enumerate(raw_rows):
        text = raw["text"]
        if index == 0:
            raw["node_type"] = "account"
        elif text.strip() == account_code:
            raw["node_type"] = "account"
        elif re.match(r"^обороты\s+за\b", text, re.IGNORECASE):
            raw["node_type"] = "month"

    # Account name/code; normalize division codes.
    for raw in raw_rows:
        if raw["node_type"] == "account":
            raw["name"] = account_code
            raw["code"] = account_code
        elif raw["node_type"] == "division":
            raw["code"] = raw["text"]
            raw["name"] = raw["text"]
        else:
            raw["code"] = None
            raw["name"] = raw["text"] or "Без названия"

    # 1С prints "<...>" when a grouping value is not resolved.
    for raw in raw_rows:
        if raw["name"].strip() in {"<...>", "<…>", "...", "…"}:
            raw["name"] = "Без названия"

    # An item row repeating the account total is the grand total line.
    account_row = next(
        (raw for raw in raw_rows if raw["node_type"] == "account"), raw_rows[0]
    )
    account_turnovers = account_row["amounts"][2:4]
    if any(account_turnovers):
        for raw in raw_rows:
            if (
                raw is not account_row
                and raw["node_type"] == "item"
                and raw["amounts"][2:4] == account_turnovers
            ):
                raw["node_type"] = "total"

    indents = [raw["indent"] for raw in raw_rows]
    min_indent = min(indents)
    distinct = len({indent - min_indent for indent in indents})

    rows: list[ParsedRow] = []

    def build_tree(levels: list[int]) -> None:
        stack: list[tuple[int, int]] = []  # (level, index in rows)
        for position, raw in enumerate(raw_rows):
            level = levels[position]
            while stack and stack[-1][0] >= level:
                stack.pop()
            parent_index = stack[-1][1] if stack else None
            month = None
            if raw["node_type"] == "month":
                month = re.sub(
                    r"^обороты\s+за\s+", "", raw["text"], flags=re.IGNORECASE
                ).strip()
            rows.append(
                ParsedRow(
                    position=position,
                    level=level,
                    node_type=raw["node_type"],
                    code=raw["code"],
                    name=raw["name"],
                    month=month or None,
                    amounts=raw["amounts"],
                    parent_index=parent_index,
                )
            )
            stack.append((level, position))

    if distinct > 1:
        levels = [raw["indent"] - min_indent for raw in raw_rows]
    else:
        # File has no indentation: derive hierarchy from node types.
        levels = [_NODE_LEVELS.get(raw["node_type"], 2) for raw in raw_rows]
    # A grand total repeats the account total and belongs to the account level,
    # not to whichever division happens to precede it.
    for index, raw in enumerate(raw_rows):
        if raw["node_type"] == "total":
            levels[index] = 1
    build_tree(levels)

    totals = account_row["amounts"] if account_row else rows[0].amounts

    return ParsedReport(
        account_code=account_code,
        company=company,
        period=period,
        source=source,
        sheet_name=ws.title,
        rows=rows,
        totals=totals,
    )
