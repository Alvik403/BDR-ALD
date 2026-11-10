from __future__ import annotations

from collections import defaultdict
from decimal import Decimal
from types import SimpleNamespace
from typing import Iterable

from .mapping import LABOR_ARTICLES, normalize_key

ZERO = Decimal("0.00")

MONTHS = {
    "январь": 1,
    "февраль": 2,
    "март": 3,
    "апрель": 4,
    "май": 5,
    "июнь": 6,
    "июль": 7,
    "август": 8,
    "сентябрь": 9,
    "октябрь": 10,
    "ноябрь": 11,
    "декабрь": 12,
}

COLUMNS = (
    {"key": "y2025", "label": "Всего 2025", "sub": "ФАКТ"},
    {"key": "m1", "label": "Январь", "sub": "ФАКТ"},
    {"key": "m2", "label": "Февраль", "sub": "ФАКТ"},
    {"key": "m3", "label": "Март", "sub": "ФАКТ"},
    {"key": "q1", "label": "1 квартал", "sub": "ФАКТ"},
    {"key": "m4", "label": "Апрель", "sub": "ФАКТ"},
    {"key": "m5", "label": "Май", "sub": "ФАКТ"},
    {"key": "m6", "label": "Июнь", "sub": "ФАКТ"},
    {"key": "h1", "label": "1 полугодие", "sub": "ФАКТ"},
    {"key": "m7", "label": "Июль", "sub": "ФАКТ"},
    {"key": "m8", "label": "Август", "sub": "ФАКТ"},
    {"key": "m9", "label": "Сентябрь", "sub": "ФАКТ"},
    {"key": "q3", "label": "3 квартал 2026", "sub": "ФАКТ"},
    {"key": "nine", "label": "9 месяцев", "sub": "ФАКТ"},
    {"key": "m10", "label": "Октябрь", "sub": "ПЛАН"},
    {"key": "m11", "label": "Ноябрь", "sub": "ПЛАН"},
    {"key": "m12", "label": "Декабрь", "sub": "ПЛАН"},
    {"key": "q4", "label": "4 квартал 2026 года", "sub": "ПЛАН"},
    {"key": "y2026", "label": "2026", "sub": "ПЛАН"},
    {"key": "y2027", "label": "2027", "sub": "ПЛАН"},
    {"key": "y2028", "label": "2028", "sub": "ПЛАН"},
    {"key": "total", "label": "Итого 2025 - 2028", "sub": "ПЛАН"},
)

REVENUE_KEYS = (
    "rev_gp",
    "rev_pir",
    "rev_smr",
    "rev_land",
    "rev_outstaffing",
    "rev_other",
)
COST_KEYS = (
    "cost_pir",
    "cost_smr",
    "cost_land",
    "cost_outstaffing",
    "cost_other",
)

MONTHLY_KEYS = {
    "rev_gp",
    "rev_pir",
    "rev_smr",
    "rev_land",
    "rev_outstaffing",
    "rev_other",
    "cost_pir",
    "cost_smr",
    "cost_land",
    "cost_outstaffing",
    "cost_other",
    "fot",
    "opex26",
    "opex25",
    "interest_income",
    "interest_expense",
    "land_interest",
    "other_income",
    "other_expense",
}

REVENUE_TARGETS = {
    "гп": "rev_gp",
    "генподрядные": "rev_gp",
    "пир": "rev_pir",
    "проектно-изыскательские работы": "rev_pir",
    "смр": "rev_smr",
    "строительно-монтажные работы": "rev_smr",
    "зу": "rev_land",
    "земля": "rev_land",
    "аутстаффинг": "rev_outstaffing",
    "прочее": "rev_other",
    "прочие доходы": "rev_other",
}
COST_TARGETS = {
    "пир": "cost_pir",
    "проектно-изыскательские работы": "cost_pir",
    "смр": "cost_smr",
    "строительно-монтажные работы": "cost_smr",
    "зу": "cost_land",
    "земля": "cost_land",
    "аутстаффинг": "cost_outstaffing",
    "прочее": "cost_other",
    "прочие расходы": "cost_other",
}
INCOME_TARGETS = {
    "% к получению": "interest_income",
    "прочие доходы": "other_income",
}
EXPENSE_TARGETS = {
    "% к уплате": "interest_expense",
    "% за зу": "land_interest",
    "прочие расходы": "other_expense",
}


def _row(section: str, param: str, key: str | None, style: str):
    return SimpleNamespace(section=section, param=param, key=key, style=style)


BDR_ROWS = (
    _row("Выручка", "", "revenue_total", "section-total"),
    _row("Основная деятельность", "Генподрядные", "rev_gp", "detail"),
    _row("Основная деятельность", "Проектно-изыскательские работы", "rev_pir", "detail"),
    _row("Основная деятельность", "Строительно-монтажные работы", "rev_smr", "detail"),
    _row("В том числе:", "Форт Сталин", "zero", "sub"),
    _row("Основная деятельность", "Земля", "rev_land", "detail"),
    _row("Основная деятельность", "ТХ", "zero", "sub"),
    _row("В том числе:", "монтаж технологического оборудования в рамках ИТП Синергия 10 11 14", "zero", "sub"),
    _row("Дополнительные доходы", "Прочие доходы", "rev_other", "detail"),
    _row("Дополнительные доходы", "17.12", "zero", "sub"),
    _row("Дополнительные доходы", "Аренда лодок", "zero", "sub"),
    _row("Дополнительные доходы", "ГОЗ (Зименс, Аутель, Альбатрос)", "zero", "sub"),
    _row("Дополнительные доходы", "БИМ Машинери", "zero", "sub"),
    _row("Дополнительные доходы", "Аренда", "zero", "sub"),
    _row("Дополнительные доходы", "Аутстаффинг", "rev_outstaffing", "detail"),
    _row("Дополнительные доходы", "Цемент", "zero", "sub"),
    _row("Дополнительные доходы", "Реализация товаров", "zero", "sub"),
    _row("Себестоимость", "", "cost_total", "section-total"),
    _row("Основная деятельность", "Проектно-изыскательские работы", "cost_pir", "detail"),
    _row("Основная деятельность", "Строительно-монтажные работы", "cost_smr", "detail"),
    _row("В том числе:", "Форт Сталин", "zero", "sub"),
    _row("Основная деятельность", "Земля", "cost_land", "detail"),
    _row("Основная деятельность", "ТХ", "zero", "sub"),
    _row("В том числе:", "монтаж технологического оборудования в рамках ИТП Синергия 10 11 14", "zero", "sub"),
    _row("Дополнительные расходы", "Прочие расходы", "cost_other", "detail"),
    _row("Дополнительные расходы", "17.12", "zero", "sub"),
    _row("Дополнительные расходы", "Аренда лодок", "zero", "sub"),
    _row("Дополнительные расходы", "ГОЗ (Зименс, Аутель, Альбатрос)", "zero", "sub"),
    _row("Дополнительные расходы", "БИМ Машинери", "zero", "sub"),
    _row("Дополнительные расходы", "Аренда", "zero", "sub"),
    _row("Дополнительные расходы", "Аутстаффинг", "cost_outstaffing", "detail"),
    _row("Дополнительные расходы", "Цемент", "zero", "sub"),
    _row("Дополнительные расходы", "Реализация товаров", "zero", "sub"),
    _row("Валовая прибыль", "", "gross_total", "section-total"),
    _row("Основная деятельность", "Генподрядные", "gross_gp", "detail"),
    _row("Основная деятельность", "Проектно-изыскательские работы", "gross_pir", "detail"),
    _row("Основная деятельность", "Строительно-монтажные работы", "gross_smr", "detail"),
    _row("В том числе:", "Форт Сталин", "zero", "sub"),
    _row("Основная деятельность", "Земля", "gross_land", "detail"),
    _row("Основная деятельность", "ТХ", "zero", "sub"),
    _row("В том числе:", "монтаж технологического оборудования в рамках ИТП Синергия 10 11 14", "zero", "sub"),
    _row("Дополнительно", "Прочее", "gross_other", "detail"),
    _row("Дополнительно", "17.12", "zero", "sub"),
    _row("Дополнительно", "Аренда лодок", "zero", "sub"),
    _row("Дополнительно", "ГОЗ (Зименс, Аутель, Альбатрос)", "zero", "sub"),
    _row("Дополнительно", "БИМ Машинери", "zero", "sub"),
    _row("Дополнительно", "Аренда", "zero", "sub"),
    _row("Дополнительно", "Аутстаффинг", "gross_outstaffing", "detail"),
    _row("Дополнительно", "Цемент", "zero", "sub"),
    _row("Дополнительно", "Реализация товаров", "zero", "sub"),
    _row("ФОТ", "", "fot", "total"),
    _row("Расходы", "Общехозяйственные расходы", "opex26", "detail"),
    _row("Расходы", "Общепроизводственные расходы", "opex25", "memo"),
    _row("Прибыль от продаж", "", "profit_sales", "total"),
    _row("Финансовая деятельность", "% к получению", "interest_income", "detail"),
    _row("Финансовая деятельность", "% к уплате", "interest_expense", "detail"),
    _row("Финансовая деятельность", "% за ЗУ", "land_interest", "detail"),
    _row("Прочие", "Прочие доходы", "other_income", "detail"),
    _row("Прочие", "Прочие расходы", "other_expense", "detail"),
    _row("Прибыль", "", None, "section"),
    _row("Прибыль до НО", "НУ", "blank", "memo"),
    _row("Прибыль до НО", "БУ", "profit_before_tax", "total"),
    _row("Налог на прибыль", "", None, "section"),
    _row("Налог на прибыль", "Отложенный налог", "blank", "memo"),
    _row("Налог на прибыль", "Текущий налог на прибыль", "blank", "memo"),
    _row("Прибыль", "", None, "section"),
    _row("Прибыль", "НУ", "blank", "memo"),
    _row("Прибыль", "БУ", "blank", "memo"),
    _row("Расчет авансовых платежей и доплат по налогу на прибыль", "", None, "section"),
    _row("Монтаж технологического оборудования в рамках ИТП Синергия 10 11 14", "", None, "section"),
    _row("", "в том числе АлДи", "blank", "memo"),
    _row("", "в том числе АО", "blank", "memo"),
)


def month_number(value: str | None) -> int | None:
    if not value:
        return None
    name = str(value).strip().split()[0].casefold()
    return MONTHS.get(name)


def _decimal(value) -> Decimal:
    return value if isinstance(value, Decimal) else Decimal(value or 0)


def aggregate(mapped_rows: Iterable) -> dict[str, dict[int, Decimal]]:
    agg: dict[str, dict[int, Decimal]] = defaultdict(
        lambda: defaultdict(Decimal)
    )
    for row in mapped_rows:
        account = getattr(row, "account_code", "")
        article = normalize_key(getattr(row, "article", None))
        is_fot = account in ("90.02", "26", "25") and article in LABOR_ARTICLES
        if getattr(row, "status", "mapped") != "mapped" and not (
            is_fot and account in ("90.02", "26")
        ):
            continue
        month = month_number(getattr(row, "month", None))
        if month is None:
            continue
        target = normalize_key(getattr(row, "target", None))
        debit = _decimal(getattr(row, "debit", ZERO))
        credit = _decimal(getattr(row, "credit", ZERO))

        if account == "90.01":
            key = REVENUE_TARGETS.get(target, "rev_other")
            agg[key][month] += credit
        elif account == "90.02":
            if article in LABOR_ARTICLES:
                agg["fot"][month] += debit
            else:
                key = COST_TARGETS.get(target, "cost_other")
                agg[key][month] += debit
        elif account == "26":
            if article in LABOR_ARTICLES:
                agg["fot"][month] += debit
            else:
                agg["opex26"][month] += debit
        elif account == "25":
            if article in LABOR_ARTICLES:
                continue
            agg["opex25"][month] += debit
        elif account == "91.01":
            key = INCOME_TARGETS.get(target, "other_income")
            agg[key][month] += credit
        elif account == "91.02":
            key = EXPENSE_TARGETS.get(target, "other_expense")
            agg[key][month] += debit
    return agg


def _monthly_values(agg: dict[str, dict[int, Decimal]]) -> dict[str, dict[int, Decimal] | None]:
    months = range(1, 13)

    def get(key: str) -> dict[int, Decimal]:
        return {month: agg.get(key, {}).get(month, ZERO) for month in months}

    values: dict[str, dict[int, Decimal] | None] = {
        key: get(key) for key in MONTHLY_KEYS
    }

    def sum_keys(keys: Iterable[str]) -> dict[int, Decimal]:
        return {
            month: sum((get(key)[month] for key in keys), ZERO)
            for month in months
        }

    def subtract(
        first: dict[int, Decimal], second: dict[int, Decimal]
    ) -> dict[int, Decimal]:
        return {month: first[month] - second[month] for month in months}

    revenue_total = sum_keys(REVENUE_KEYS)
    cost_total = sum_keys(COST_KEYS)
    values["revenue_total"] = revenue_total
    values["cost_total"] = cost_total
    values["gross_total"] = subtract(revenue_total, cost_total)

    values["gross_gp"] = get("rev_gp")
    values["gross_pir"] = subtract(get("rev_pir"), get("cost_pir"))
    values["gross_smr"] = subtract(get("rev_smr"), get("cost_smr"))
    values["gross_land"] = subtract(get("rev_land"), get("cost_land"))
    values["gross_other"] = subtract(get("rev_other"), get("cost_other"))
    values["gross_outstaffing"] = subtract(
        get("rev_outstaffing"), get("cost_outstaffing")
    )

    fot = get("fot")
    opex26 = get("opex26")
    values["fot"] = fot
    values["opex26"] = opex26
    values["opex25"] = get("opex25")
    profit_sales = subtract(
        subtract(values["gross_total"], fot), opex26
    )
    values["profit_sales"] = profit_sales
    profit_before_tax = {
        month: (
            profit_sales[month]
            + get("interest_income")[month]
            - get("interest_expense")[month]
            - get("land_interest")[month]
            + get("other_income")[month]
            - get("other_expense")[month]
        )
        for month in months
    }
    values["profit_before_tax"] = profit_before_tax
    values["blank"] = None
    values["zero"] = {month: ZERO for month in months}
    return values


def _column_value(
    monthly: dict[int, Decimal] | None,
    column: str,
    available_months: set[int],
) -> Decimal | None:
    if monthly is None:
        return None
    if column == "y2025":
        return None
    if column.startswith("m") and column[1:].isdigit():
        month = int(column[1:])
        if month not in available_months:
            return None
        return monthly.get(month, ZERO)

    groups = {
        "q1": {1, 2, 3},
        "h1": {1, 2, 3, 4, 5, 6},
        "nine": set(range(1, 10)),
        "q3": {7, 8, 9},
    }
    if column in groups:
        months = groups[column]
        if not months.issubset(available_months):
            return None
        return sum((monthly.get(month, ZERO) for month in months), ZERO)

    # План 4 квартала, годовые и итог пока не имеют источника данных.
    return None


def build_bdr(
    mapped_rows: Iterable,
    available_months: set[int] | None = None,
) -> dict:
    rows = list(mapped_rows)
    if available_months is None:
        available_months = {
            month
            for month in (month_number(getattr(row, "month", None)) for row in rows)
            if month is not None
        }

    agg = aggregate(rows)
    values = _monthly_values(agg)
    result_rows = []
    for spec in BDR_ROWS:
        monthly = values.get(spec.key) if spec.key else None
        row_values = {
            column["key"]: _column_value(
                monthly, column["key"], available_months
            )
            for column in COLUMNS
        }
        result_rows.append(
            SimpleNamespace(
                section=spec.section,
                param=spec.param,
                style=spec.style,
                values=row_values,
            )
        )

    def summary_value(key: str) -> Decimal | None:
        monthly = values.get(key)
        if monthly is None:
            return None
        if not {7, 8, 9}.issubset(available_months):
            return None
        return sum((monthly.get(month, ZERO) for month in (7, 8, 9)), ZERO)

    return {
        "columns": COLUMNS,
        "rows": result_rows,
        "available_months": sorted(available_months),
        "has_unmapped": any(
            getattr(row, "status", "mapped") == "unmapped"
            and not (
                getattr(row, "account_code", "") in ("90.02", "26")
                and normalize_key(getattr(row, "article", None))
                in LABOR_ARTICLES
            )
            and (
                _decimal(getattr(row, "debit", ZERO)) != ZERO
                or _decimal(getattr(row, "credit", ZERO)) != ZERO
            )
            for row in rows
        ),
        "summary": {
            "revenue": summary_value("revenue_total"),
            "cost": summary_value("cost_total"),
            "gross": summary_value("gross_total"),
            "fot": summary_value("fot"),
            "profit_sales": summary_value("profit_sales"),
            "profit_before_tax": summary_value("profit_before_tax"),
        },
    }
