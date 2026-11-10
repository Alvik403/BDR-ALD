from __future__ import annotations

import hashlib
import logging
import os
from contextlib import asynccontextmanager
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from types import SimpleNamespace
from urllib.parse import urlencode

from fastapi import Depends, FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import (
    FileResponse,
    HTMLResponse,
    JSONResponse,
    RedirectResponse,
    Response,
)
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlalchemy import func, select, text
from sqlalchemy.orm import Session

from .bdr import build_bdr, month_number
from .db import Base, SessionLocal, engine, get_db, wait_for_db
from .mapping import (
    ALLOWED_MAPPING_KINDS,
    LABOR_ARTICLES,
    MappingImportError,
    apply_mapping_entries,
    build_fot_mapping_export,
    build_fot_mapping_template,
    build_mapping_export,
    build_mapping_template,
    extract_year,
    normalize_key,
    parse_mapping_workbook,
    rebuild_mapped_rows,
)
from .models import Account, MappedRow, MappingEntry, OsvRow, Upload
from .parser import ParseError, parse_osv

log = logging.getLogger("osv")
logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = Path(os.getenv("UPLOAD_DIR", "data/uploads")).resolve()
MAX_UPLOAD_BYTES = int(os.getenv("MAX_UPLOAD_MB", "25")) * 1024 * 1024
ALLOWED_EXTENSIONS = {".xlsx", ".xlsm"}

SEED_ACCOUNTS = [
    (
        "25",
        "Общепроизводственные расходы",
        "Счёт 25 · подразделения, статьи затрат, помесячная расшифровка. "
        "Импорт проверен на файле 25_ОСВ.xlsx.",
        10,
        True,
    ),
    (
        "26",
        "Общехозяйственные расходы",
        "Счёт 26 · тот же универсальный разбор ОСВ, что и для счёта 25.",
        20,
        False,
    ),
    (
        "90.01",
        "Выручка",
        "Счёт 90.01 · выручка от продаж.",
        30,
        False,
    ),
    (
        "90.02",
        "Себестоимость продаж",
        "Счёт 90.02 · себестоимость продаж.",
        40,
        False,
    ),
    (
        "91.01",
        "Прочие доходы",
        "Счёт 91.01 · прочие доходы.",
        50,
        False,
    ),
    (
        "91.02",
        "Прочие расходы",
        "Счёт 91.02 · прочие расходы.",
        60,
        False,
    ),
]

templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))
ASSET_VERSION = "20261103-12"
templates.env.globals["asset_version"] = ASSET_VERSION

ZERO = Decimal("0.00")

UNIFIED_KIND_LABELS = {
    "article": "Статья",
    "nomenclature": "Номенклатурная группа",
    "department": "Подразделение",
}

FOT_ACCOUNTS = ("90.02", "26", "25")
FOT_ACCOUNT_TITLES = {
    "90.02": "90.02 · себестоимость",
    "26": "26 · общехозяйственные расходы",
    "25": "25 · общепроизводственные расходы",
}
FOT_CATEGORY_NONE = "Без категории"


def _money(value, zero: str = "—") -> str:
    if value is None:
        return zero
    amount = Decimal(value)
    if amount == 0:
        return zero
    return f"{amount:,.2f}".replace(",", "\u00a0").replace(".", ",")


def _money0(value) -> str:
    return _money(value, zero="0,00")


def _dt(value: datetime | None) -> str:
    if value is None:
        return "—"
    return value.strftime("%d.%m.%Y %H:%M")


def _num(value) -> str:
    if value is None:
        return "—"
    try:
        return f"{int(value):,}".replace(",", "\u00a0")
    except (TypeError, ValueError):
        return "—"


templates.env.filters["money"] = _money
templates.env.filters["money0"] = _money0
templates.env.filters["dt"] = _dt
templates.env.filters["num"] = _num

GENERAL_PAGE_SIZE = 200
MONTH_ORDER = {
    name: index
    for index, names in enumerate(
        (
            ("январь", "января"),
            ("февраль", "февраля"),
            ("март", "марта"),
            ("апрель", "апреля"),
            ("май", "мая"),
            ("июнь", "июня"),
            ("июль", "июля"),
            ("август", "августа"),
            ("сентябрь", "сентября"),
            ("октябрь", "октября"),
            ("ноябрь", "ноября"),
            ("декабрь", "декабря"),
        ),
        start=1,
    )
    for name in names
}


def _month_sort_key(value: str | None) -> tuple[int, int, str]:
    if not value:
        return (9999, 99, "")
    parts = value.split()
    name = parts[0].casefold() if parts else ""
    year = 0
    for part in parts[1:]:
        digits = "".join(char for char in part if char.isdigit())
        if digits:
            year = int(digits)
            break
    return (year, MONTH_ORDER.get(name, 99), value)


def _decimal_filter(value: str | None) -> Decimal | None:
    if not value:
        return None
    text = value.replace("\xa0", "").replace(" ", "").replace(",", ".")
    try:
        return Decimal(text)
    except Exception:  # noqa: BLE001 - invalid user input is ignored
        return None


def _display_name(row: OsvRow) -> str:
    if row.node_type == "account":
        return f"Счёт {row.code or row.name}"
    if row.node_type == "division":
        return f"Подразделение {row.code or row.name}"
    if row.node_type == "total":
        return "Итого"
    return row.name


def _redirect(path: str, **params) -> RedirectResponse:
    fragment = str(params.pop("_fragment", "") or "")
    query = urlencode({key: value for key, value in params.items() if value})
    url = path + (f"?{query}" if query else "")
    if fragment:
        url += f"#{fragment}"
    return RedirectResponse(url, status_code=303)


async def _seed(db: Session) -> None:
    for code, name, description, order, ready in SEED_ACCOUNTS:
        account = db.get(Account, code)
        if account is None:
            db.add(
                Account(
                    code=code,
                    name=name,
                    description=description,
                    sort_order=order,
                    is_ready=ready,
                )
            )
        else:
            account.name = name
            account.description = description
            account.sort_order = order
            account.is_ready = ready
    db.commit()


def _migrate_mapping_schema() -> None:
    with engine.begin() as conn:
        conn.execute(
            text(
                "ALTER TABLE IF EXISTS mapping_entries "
                "ADD COLUMN IF NOT EXISTS account_code "
                "VARCHAR(20) NOT NULL DEFAULT ''"
            )
        )
        conn.execute(
            text(
                "ALTER TABLE IF EXISTS mapping_entries "
                "ADD COLUMN IF NOT EXISTS year "
                "INTEGER NOT NULL DEFAULT 0"
            )
        )
        conn.execute(
            text(
                "ALTER TABLE IF EXISTS mapping_entries "
                "DROP CONSTRAINT IF EXISTS uq_mapping_kind_key"
            )
        )
        conn.execute(
            text(
                "CREATE UNIQUE INDEX IF NOT EXISTS "
                "uq_mapping_kind_account_year_key "
                "ON mapping_entries "
                "(kind, account_code, year, source_key)"
            )
        )


@asynccontextmanager
async def lifespan(app: FastAPI):
    wait_for_db()
    Base.metadata.create_all(engine)
    _migrate_mapping_schema()
    with SessionLocal() as db:
        await _seed(db)
        try:
            mapped = rebuild_mapped_rows(db)
            log.info("Mapping rebuilt for %s rows", mapped)
        except Exception:  # noqa: BLE001 - do not block startup
            log.exception("Failed to rebuild mapping on startup")
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    log.info("OSV application started, uploads in %s", UPLOAD_DIR)
    yield


app = FastAPI(title="ОСВ · учёт", lifespan=lifespan)
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")


def _latest_upload(db: Session, account_code: str) -> Upload | None:
    return db.scalars(
        select(Upload)
        .where(Upload.account_code == account_code)
        .order_by(Upload.created_at.desc())
    ).first()


def _account_tabs(db: Session, current: str) -> list[SimpleNamespace]:
    accounts = db.scalars(
        select(Account).order_by(Account.sort_order, Account.code)
    ).all()
    counts = dict(
        db.execute(
            select(Upload.account_code, func.count(Upload.id)).group_by(
                Upload.account_code
            )
        ).all()
    )
    return [
        SimpleNamespace(
            code=account.code,
            name=account.name,
            has_data=counts.get(account.code, 0) > 0,
            is_current=account.code == current,
        )
        for account in accounts
    ]


def _effective_mapping_entries(
    entries: list[MappingEntry],
    account_code: str,
    target_year: int,
) -> dict[str, MappingEntry]:
    candidates = [account_code]
    if "." in account_code:
        candidates.append(account_code.split(".", 1)[0])
    candidates.append("")
    account_rank = {
        candidate: index
        for index, candidate in enumerate(dict.fromkeys(candidates))
    }
    best: dict[str, tuple[tuple[int, int], MappingEntry]] = {}
    for entry in entries:
        if entry.account_code not in account_rank:
            continue
        entry_year = int(entry.year or 0)
        if entry_year not in (target_year, 0):
            continue
        target = (entry.target or "").strip()
        if not target:
            continue
        rank = (
            account_rank[entry.account_code],
            0 if entry_year == target_year else 1,
        )
        current = best.get(entry.source_key)
        if current is None or rank < current[0]:
            best[entry.source_key] = (rank, entry)
    return {
        key: entry for key, (_rank, entry) in best.items()
    }


def _month_summary(rows: list[OsvRow]) -> list[dict]:
    months: list[dict] = []
    index: dict[str, dict] = {}
    for row in rows:
        if row.node_type != "month":
            continue
        key = row.month or "Без месяца"
        bucket = index.get(key)
        if bucket is None:
            bucket = {"name": key, "debit": ZERO, "credit": ZERO}
            index[key] = bucket
            months.append(bucket)
        bucket["debit"] += row.turnover_debit
        bucket["credit"] += row.turnover_credit
    return months


@app.get("/", response_class=HTMLResponse)
def index(request: Request, db: Session = Depends(get_db)):
    accounts = db.scalars(
        select(Account).order_by(Account.sort_order, Account.code)
    ).all()
    cards = [
        SimpleNamespace(account=account, upload=_latest_upload(db, account.code))
        for account in accounts
    ]
    loaded = [item.upload for item in cards if item.upload is not None]
    periods = {upload.period for upload in loaded if upload.period}
    return templates.TemplateResponse(
        request,
        "index.html",
        {
            "active_nav": "dashboard",
            "cards": cards,
            "summary": {
                "accounts": len(accounts),
                "with_data": len(loaded),
                "rows": sum(upload.row_count for upload in loaded),
                "last": max(
                    (upload.created_at for upload in loaded),
                    default=None,
                ),
            },
            "period": next(iter(periods)) if len(periods) == 1 else None,
            **_fot_overview(db),
        },
    )


@app.get("/accounts/{code}", response_class=HTMLResponse)
def account_page(
    code: str,
    request: Request,
    upload_id: str | None = None,
    view: str = "grouped",
    mapping_filter: str = "",
    ok: str | None = None,
    err: str | None = None,
    db: Session = Depends(get_db),
):
    account = db.get(Account, code)
    if account is None:
        raise HTTPException(status_code=404, detail="Счёт не найден")

    uploads = db.scalars(
        select(Upload)
        .where(Upload.account_code == code)
        .order_by(Upload.created_at.desc())
    ).all()
    upload = next((item for item in uploads if item.id == upload_id), None)
    if upload is None and uploads:
        upload = uploads[0]
    view = "flat" if view == "flat" else "grouped"
    mapping_filter = "unmapped" if mapping_filter == "unmapped" else ""

    rows: list[OsvRow] = []
    row_views: list[SimpleNamespace] = []
    months: list[dict] = []
    flat_rows: list[SimpleNamespace] = []
    flat_totals: tuple[Decimal, ...] = (ZERO, ZERO, ZERO, ZERO, ZERO, ZERO)
    flat_has_balances = False
    flat_has_months = False
    flat_has_divisions = False
    flat_has_groups = False

    mapped_rows: list[MappedRow] = []
    mapping_mapped_count = 0
    mapping_unmapped_count = 0
    mapping_has_department = False
    mapping_has_group = False
    mapping_has_article = False
    mapping_total_debit = ZERO
    mapping_total_credit = ZERO

    if upload is not None:
        rows = db.scalars(
            select(OsvRow)
            .where(OsvRow.upload_id == upload.id)
            .order_by(OsvRow.position)
        ).all()
        parent_ids = {row.parent_id for row in rows if row.parent_id is not None}
        month_parents = {
            row.parent_id for row in rows if row.node_type == "month"
        }
        row_views = [
            SimpleNamespace(
                row=row,
                display=_display_name(row),
                has_children=row.id in parent_ids,
                has_month_children=row.id in month_parents,
                is_group=row.id in parent_ids and row.id not in month_parents,
            )
            for row in rows
        ]
        months = _month_summary(rows)

        by_id = {row.id: row for row in rows}

        running = [ZERO] * 6
        for row in rows:
            if row.id in parent_ids or row.node_type in (
                "account",
                "division",
                "total",
            ):
                continue

            ancestors: list[OsvRow] = []
            parent = by_id.get(row.parent_id)
            while parent is not None:
                if parent.node_type != "account":
                    ancestors.append(parent)
                parent = by_id.get(parent.parent_id)
            ancestors.reverse()

            if row.node_type == "month":
                item = ancestors[-1] if ancestors else None
                grouping = ancestors[:-1]
            else:
                item = row
                grouping = ancestors

            division = grouping[0] if grouping else None
            middle = " / ".join(
                node.code or node.name for node in grouping[1:]
            )

            flat_rows.append(
                SimpleNamespace(
                    division=(
                        (division.code or division.name) if division else ""
                    ),
                    group=middle,
                    item=item.name if item else row.name,
                    month=row.month or "",
                    row=row,
                )
            )
            amounts = (
                row.opening_debit,
                row.opening_credit,
                row.turnover_debit,
                row.turnover_credit,
                row.closing_debit,
                row.closing_credit,
            )
            for index in range(6):
                running[index] += amounts[index]

        flat_totals = tuple(running)
        flat_has_balances = any(
            item.row.opening_debit
            or item.row.opening_credit
            or item.row.closing_debit
            or item.row.closing_credit
            for item in flat_rows
        )
        flat_has_months = any(item.month for item in flat_rows)
        flat_has_divisions = any(item.division for item in flat_rows)
        flat_has_groups = any(item.group for item in flat_rows)

        mapped_rows = db.scalars(
            select(MappedRow)
            .where(MappedRow.upload_id == upload.id)
            .order_by(MappedRow.position)
        ).all()
        mapping_mapped_count = sum(
            1 for row in mapped_rows if row.status == "mapped"
        )
        mapping_unmapped_count = len(mapped_rows) - mapping_mapped_count
        if mapping_filter == "unmapped":
            mapped_rows = [
                row for row in mapped_rows if row.status == "unmapped"
            ]
        mapping_has_department = any(
            row.mapping_kind == "department" for row in mapped_rows
        )
        mapping_has_group = any(row.group_name for row in mapped_rows)
        mapping_has_article = any(row.article for row in mapped_rows)
        mapping_total_debit = sum((row.debit for row in mapped_rows), ZERO)
        mapping_total_credit = sum((row.credit for row in mapped_rows), ZERO)

    flat_meta_cols = (
        1 + int(flat_has_divisions) + int(flat_has_groups) + int(flat_has_months)
    )

    return templates.TemplateResponse(
        request,
        "account.html",
        {
            "active_nav": "account",
            "account": account,
            "account_tabs": _account_tabs(db, code),
            "uploads": uploads,
            "upload": upload,
            "rows": row_views,
            "months": months,
            "view": view,
            "flat_rows": flat_rows,
            "flat_totals": flat_totals,
            "flat_has_balances": flat_has_balances,
            "flat_has_months": flat_has_months,
            "flat_has_divisions": flat_has_divisions,
            "flat_has_groups": flat_has_groups,
            "flat_meta_cols": flat_meta_cols,
            "mapped_rows": mapped_rows,
            "mapping_mapped_count": mapping_mapped_count,
            "mapping_unmapped_count": mapping_unmapped_count,
            "mapping_has_department": mapping_has_department,
            "mapping_has_group": mapping_has_group,
            "mapping_has_article": mapping_has_article,
            "mapping_total_debit": mapping_total_debit,
            "mapping_total_credit": mapping_total_credit,
            "mapping_filter": mapping_filter,
            "ok": ok,
            "err": err,
        },
    )


@app.get("/general", response_class=HTMLResponse)
def general(
    request: Request,
    account: str = "",
    department: str = "",
    group: str = "",
    article: str = "",
    month: str = "",
    target: str = "",
    status: str = "",
    debit_min: str = "",
    debit_max: str = "",
    credit_min: str = "",
    credit_max: str = "",
    page: int = 1,
    db: Session = Depends(get_db),
):
    accounts = db.scalars(
        select(Account).order_by(Account.sort_order, Account.code)
    ).all()
    latest_ids = db.scalars(
        select(Upload.id)
        .distinct(Upload.account_code)
        .order_by(Upload.account_code, Upload.created_at.desc())
    ).all()

    conditions = []
    if latest_ids:
        conditions.append(MappedRow.upload_id.in_(latest_ids))
    else:
        conditions.append(MappedRow.id < 0)

    if account:
        conditions.append(MappedRow.account_code == account)
    if department:
        conditions.append(
            MappedRow.department.ilike(f"%{department.strip()}%")
        )
    if group:
        conditions.append(MappedRow.group_name.ilike(f"%{group.strip()}%"))
    if article:
        conditions.append(MappedRow.article.ilike(f"%{article.strip()}%"))
    if month:
        conditions.append(MappedRow.month == month)
    if target:
        conditions.append(MappedRow.target.ilike(f"%{target.strip()}%"))
    if status == "unmapped":
        conditions.append(MappedRow.status == "unmapped")
    elif status == "mapped":
        conditions.append(MappedRow.status == "mapped")

    debit_min_value = _decimal_filter(debit_min)
    debit_max_value = _decimal_filter(debit_max)
    credit_min_value = _decimal_filter(credit_min)
    credit_max_value = _decimal_filter(credit_max)
    if debit_min_value is not None:
        conditions.append(MappedRow.debit >= debit_min_value)
    if debit_max_value is not None:
        conditions.append(MappedRow.debit <= debit_max_value)
    if credit_min_value is not None:
        conditions.append(MappedRow.credit >= credit_min_value)
    if credit_max_value is not None:
        conditions.append(MappedRow.credit <= credit_max_value)

    summary = db.execute(
        select(
            func.count(MappedRow.id),
            func.coalesce(func.sum(MappedRow.debit), 0),
            func.coalesce(func.sum(MappedRow.credit), 0),
        ).where(*conditions)
    ).one()
    total = summary[0] or 0
    total_debit = Decimal(summary[1] or 0)
    total_credit = Decimal(summary[2] or 0)

    total_pages = max(
        1, (total + GENERAL_PAGE_SIZE - 1) // GENERAL_PAGE_SIZE
    )
    page = min(max(page, 1), total_pages)
    offset = (page - 1) * GENERAL_PAGE_SIZE

    rows = db.scalars(
        select(MappedRow)
        .where(*conditions)
        .order_by(MappedRow.account_code, MappedRow.position)
        .offset(offset)
        .limit(GENERAL_PAGE_SIZE)
    ).all()

    if latest_ids:
        month_options = sorted(
            db.scalars(
                select(MappedRow.month)
                .where(
                    MappedRow.upload_id.in_(latest_ids),
                    MappedRow.month.isnot(None),
                )
                .distinct()
            ).all(),
            key=_month_sort_key,
        )
    else:
        month_options = []

    filter_values = {
        "account": account,
        "department": department,
        "group": group,
        "article": article,
        "month": month,
        "target": target,
        "status": status,
        "debit_min": debit_min,
        "debit_max": debit_max,
        "credit_min": credit_min,
        "credit_max": credit_max,
    }
    clean_filters = {key: value for key, value in filter_values.items() if value}

    def page_url(number: int) -> str:
        params = dict(clean_filters)
        if number > 1:
            params["page"] = number
        return "/general" + (f"?{urlencode(params)}" if params else "")

    return templates.TemplateResponse(
        request,
        "general.html",
        {
            "active_nav": "general",
            "accounts": accounts,
            "rows": rows,
            "total": total,
            "total_debit": total_debit,
            "total_credit": total_credit,
            "page": page,
            "total_pages": total_pages,
            "page_size": GENERAL_PAGE_SIZE,
            "month_options": month_options,
            "filters": filter_values,
            "page_url": page_url,
        },
    )


@app.get("/bdr", response_class=HTMLResponse)
def bdr_page(request: Request, db: Session = Depends(get_db)):
    latest_ids = db.scalars(
        select(Upload.id)
        .distinct(Upload.account_code)
        .order_by(Upload.account_code, Upload.created_at.desc())
    ).all()
    rows = (
        db.scalars(
            select(MappedRow).where(
                MappedRow.upload_id.in_(latest_ids)
            )
        ).all()
        if latest_ids
        else []
    )
    available_months = {
        month
        for row in rows
        if (month := month_number(row.month)) is not None
    }
    bdr = build_bdr(rows, available_months)
    return templates.TemplateResponse(
        request,
        "bdr.html",
        {
            "active_nav": "bdr",
            "bdr": bdr,
        },
    )


def _fot_mapping_rows(
    db: Session,
    department_entries: list[MappingEntry],
    upload: Upload,
) -> list[SimpleNamespace]:
    code = upload.account_code
    year = extract_year(upload.period)
    rows = db.scalars(
        select(MappedRow).where(MappedRow.upload_id == upload.id)
    ).all()
    buckets: dict[str, dict] = {}
    for row in rows:
        if normalize_key(row.article) not in LABOR_ARTICLES:
            continue
        department = (row.department or "").strip() or "Без названия"
        key = normalize_key(department)
        buckets.setdefault(key, {"department": department, "key": key})
    effective = _effective_mapping_entries(department_entries, code, year)
    result: list[SimpleNamespace] = []
    for index, item in enumerate(
        sorted(buckets.values(), key=lambda item: item["department"].casefold()),
        start=1,
    ):
        entry = effective.get(item["key"])
        target = (entry.target or "").strip() if entry else ""
        result.append(
            SimpleNamespace(
                account=code,
                year=year,
                department=item["department"],
                target=target,
                entry_id=entry.id if entry else None,
                form_suffix=f"fot_{code.replace('.', '_')}_{index}",
            )
        )
    return result


def _fot_overview(db: Session, account: str = "") -> dict:
    """ФОТ-аналитика: суммы, категории и подразделения по 25 / 26 / 90.02."""
    account = account if account in FOT_ACCOUNTS else ""
    latest_uploads = db.scalars(
        select(Upload)
        .where(Upload.account_code.in_(FOT_ACCOUNTS))
        .distinct(Upload.account_code)
        .order_by(Upload.account_code, Upload.created_at.desc())
    ).all()
    latest_by_account = {
        upload.account_code: upload for upload in latest_uploads
    }
    latest_ids = [upload.id for upload in latest_uploads]
    entries = db.scalars(
        select(MappingEntry)
        .where(MappingEntry.kind == "department")
        .order_by(MappingEntry.source_text)
    ).all()
    account_years = {
        code: extract_year(upload.period)
        for code, upload in latest_by_account.items()
    }
    effective = {
        code: _effective_mapping_entries(
            entries, code, account_years.get(code, 0)
        )
        for code in FOT_ACCOUNTS
    }

    rows = (
        db.scalars(
            select(MappedRow)
            .where(MappedRow.upload_id.in_(latest_ids))
            .order_by(MappedRow.account_code, MappedRow.position)
        ).all()
        if latest_ids
        else []
    )

    buckets: dict[str, dict[str, dict]] = {
        code: {} for code in FOT_ACCOUNTS
    }
    for row in rows:
        code = row.account_code
        if code not in FOT_ACCOUNTS:
            continue
        if normalize_key(row.article) not in LABOR_ARTICLES:
            continue
        department = (row.department or "").strip() or "Без названия"
        key = normalize_key(department)
        bucket = buckets[code].setdefault(
            key,
            {"department": department, "key": key, "amount": ZERO},
        )
        bucket["amount"] += row.debit or ZERO

    groups = []
    category_index: dict[str, dict] = {}
    totals = {code: ZERO for code in FOT_ACCOUNTS}
    for code in FOT_ACCOUNTS:
        items = sorted(
            buckets[code].values(),
            key=lambda item: (
                -item["amount"],
                item["department"].casefold(),
            ),
        )
        total = sum((item["amount"] for item in items), ZERO)
        totals[code] = total
        for index, item in enumerate(items, start=1):
            entry = effective.get(code, {}).get(item["key"])
            target = (entry.target or "").strip() if entry else ""
            item.update(
                {
                    "target": target,
                    "entry_id": entry.id if entry else None,
                    "year": account_years.get(code, 0),
                    "form_suffix": (
                        f"department_{code.replace('.', '_')}_{index}"
                    ),
                    "status": "mapped" if target else "unmapped",
                    "account": code,
                }
            )
            category = target or FOT_CATEGORY_NONE
            category_row = category_index.setdefault(
                category,
                {
                    "target": category,
                    "by_account": {key: ZERO for key in FOT_ACCOUNTS},
                    "total_bdr": ZERO,
                },
            )
            category_row["by_account"][code] += item["amount"]
            if code in ("90.02", "26"):
                category_row["total_bdr"] += item["amount"]

        groups.append(
            SimpleNamespace(
                code=code,
                title=FOT_ACCOUNT_TITLES[code],
                rows=items,
                total=total,
                in_bdr=code != "25",
                has_unmapped=any(not item["target"] for item in items),
                note=(
                    "Справочно: ФОТ 25 счёта не входит в итог БДР."
                    if code == "25"
                    else ""
                ),
            )
        )

    category_rows = sorted(
        category_index.values(),
        key=lambda item: (
            item["target"] == FOT_CATEGORY_NONE,
            -item["total_bdr"],
            item["target"].casefold(),
        ),
    )
    unmapped_amount = sum(
        (
            item["amount"]
            for group in groups
            for item in group.rows
            if not item["target"]
        ),
        ZERO,
    )
    periods = sorted(
        {
            upload.period
            for upload in latest_by_account.values()
            if upload.period
        }
    )
    period_label = periods[0] if len(periods) == 1 else "последние загрузки"
    summary = SimpleNamespace(
        bdr=totals["90.02"] + totals["26"],
        account_9002=totals["90.02"],
        account_26=totals["26"],
        account_25=totals["25"],
        unmapped=unmapped_amount,
    )
    visible_groups = [
        group
        for group in groups
        if not account or group.code == account
    ]
    return {
        "fot_account_filter": account,
        "fot_period_label": period_label,
        "fot_summary": summary,
        "fot_category_rows": category_rows,
        "fot_groups": visible_groups,
        "fot_accounts": [
            SimpleNamespace(code=code, title=FOT_ACCOUNT_TITLES[code])
            for code in FOT_ACCOUNTS
        ],
    }


@app.get("/mapping", response_class=HTMLResponse)
def mapping_page(
    request: Request,
    view: str = "main",
    account: str = "",
    kind: str = "",
    status: str = "",
    q: str = "",
    ok: str | None = None,
    err: str | None = None,
    db: Session = Depends(get_db),
):
    view = view if view in ("main", "fot") else "main"
    accounts = db.scalars(
        select(Account).order_by(Account.sort_order, Account.code)
    ).all()
    account_codes = [item.code for item in accounts]
    if account not in account_codes:
        account = ""
    if kind not in UNIFIED_KIND_LABELS:
        kind = ""
    if status not in ("unmapped", "mapped"):
        status = ""
    query_text = (q or "").strip()
    entries = db.scalars(
        select(MappingEntry).order_by(
            MappingEntry.account_code,
            MappingEntry.kind,
            MappingEntry.source_text,
        )
    ).all()

    latest_uploads = db.scalars(
        select(Upload)
        .distinct(Upload.account_code)
        .order_by(Upload.account_code, Upload.created_at.desc())
    ).all()
    latest_by_account = {
        upload.account_code: upload for upload in latest_uploads
    }
    latest_ids = [upload.id for upload in latest_uploads]
    account_years = {
        upload.account_code: extract_year(upload.period)
        for upload in latest_uploads
    }
    scope_conditions = (
        [MappedRow.upload_id.in_(latest_ids)]
        if latest_ids
        else [MappedRow.id < 0]
    )

    unmapped_query = select(
        MappedRow.mapping_kind,
        MappedRow.account_code,
        MappedRow.mapping_key,
        func.min(MappedRow.article).label("sample_article"),
        func.min(MappedRow.group_name).label("sample_group"),
        func.min(MappedRow.department).label("sample_department"),
    ).where(MappedRow.status == "unmapped", *scope_conditions)
    unmapped_query = unmapped_query.group_by(
        MappedRow.mapping_kind,
        MappedRow.account_code,
        MappedRow.mapping_key,
    )

    unmapped = []
    for (
        row_kind,
        account_code,
        key,
        sample_article,
        sample_group,
        sample_department,
    ) in db.execute(unmapped_query).all():
        sample_text = {
            "article": sample_article,
            "nomenclature": sample_group,
            "department": sample_department,
        }.get(row_kind) or key
        unmapped.append(
            SimpleNamespace(
                kind=row_kind,
                key=key,
                text=sample_text,
                account=account_code or "",
                year=(
                    0
                    if row_kind == "article"
                    else account_years.get(account_code or "", 0)
                ),
            )
        )

    rows_by_key: dict[tuple[str, str, int, str], SimpleNamespace] = {}
    for entry in entries:
        if entry.kind == "department" and entry.account_code == "25":
            continue
        entry_year = int(entry.year or 0)
        key = (
            entry.kind,
            entry.account_code or "",
            entry_year,
            entry.source_key,
        )
        rows_by_key[key] = SimpleNamespace(
            account=entry.account_code or "",
            year=entry_year,
            kind=entry.kind,
            type_label=UNIFIED_KIND_LABELS.get(entry.kind, entry.kind),
            source_text=entry.source_text,
            target=(entry.target or "").strip(),
            entry_id=entry.id,
            form_suffix="",
        )
    for item in unmapped:
        if item.kind == "department" and item.account == "25":
            continue
        key = (item.kind, item.account, int(item.year or 0), item.key)
        if key in rows_by_key:
            continue
        rows_by_key[key] = SimpleNamespace(
            account=item.account,
            year=int(item.year or 0),
            kind=item.kind,
            type_label=UNIFIED_KIND_LABELS.get(item.kind, item.kind),
            source_text=item.text,
            target="",
            entry_id=None,
            form_suffix="",
        )

    unified_rows = list(rows_by_key.values())
    kind_order = {"article": 0, "nomenclature": 1, "department": 2}

    def row_sort_key(row: SimpleNamespace):
        return (
            1 if row.target else 0,
            row.account or "~",
            kind_order.get(row.kind, 9),
            (row.source_text or "").casefold(),
        )

    unified_rows.sort(key=row_sort_key)
    if account:
        unified_rows = [
            row for row in unified_rows if row.account == account
        ]
    if kind:
        unified_rows = [row for row in unified_rows if row.kind == kind]
    if status == "unmapped":
        unified_rows = [row for row in unified_rows if not row.target]
    elif status == "mapped":
        unified_rows = [row for row in unified_rows if row.target]
    if query_text:
        needle = query_text.casefold()
        unified_rows = [
            row
            for row in unified_rows
            if needle in (row.source_text or "").casefold()
            or needle in (row.target or "").casefold()
            or needle in row.type_label.casefold()
        ]

    for index, row in enumerate(unified_rows, start=1):
        row.form_suffix = f"unified_{index}"

    fot_rows = []
    if view == "fot":
        department_entries = [
            entry for entry in entries if entry.kind == "department"
        ]
        for code in ("26", "90.02"):
            upload = latest_by_account.get(code)
            if upload is not None:
                fot_rows.extend(
                    _fot_mapping_rows(db, department_entries, upload)
                )
        fot_rows.sort(
            key=lambda row: (
                1 if row.target else 0,
                row.account,
                row.department.casefold(),
            )
        )

    main_params: dict[str, str] = {"view": "main"}
    if view == "main":
        if account:
            main_params["account"] = account
        if kind:
            main_params["kind"] = kind
        if status:
            main_params["status"] = status
        if query_text:
            main_params["q"] = query_text
    main_url = "/mapping?" + urlencode(main_params)

    fot_url = "/mapping?view=fot"

    return templates.TemplateResponse(
        request,
        "mapping.html",
        {
            "active_nav": "mapping",
            "view": view,
            "main_url": main_url,
            "fot_url": fot_url,
            "accounts": accounts,
            "account_filter": account,
            "kind_filter": kind,
            "status_filter": status,
            "q_filter": query_text,
            "kind_options": [
                ("article", "Статья"),
                ("nomenclature", "Номенклатурная группа"),
                ("department", "Подразделение"),
            ],
            "unified_rows": unified_rows,
            "fot_rows": fot_rows,
            "ok": ok,
            "err": err,
        },
    )


@app.get("/fot")
def fot_page(
    account: str = "",
    ok: str | None = None,
    err: str | None = None,
):
    params: dict[str, str] = {"view": "fot"}
    if account:
        params["account"] = account
    if ok:
        params["ok"] = ok
    if err:
        params["err"] = err
    return _redirect("/mapping", **params)


@app.post("/mapping/save")
async def mapping_save(request: Request, db: Session = Depends(get_db)):
    form = await request.form()
    account = str(form.get("account", "")).strip()
    return_to = str(form.get("return_to", "")).strip()
    return_kind = str(form.get("return_kind", "")).strip()
    return_status = str(form.get("return_status", "")).strip()
    return_q = str(form.get("return_q", "")).strip()
    return_page = str(form.get("return_page", "")).strip()
    view = "main"
    anchor = ""
    redirect_path = "/mapping"
    if return_to in ("mapping-fot", "fot"):
        view = "fot"
        anchor = "fot"
    elif return_to.startswith("mapping-"):
        anchor = return_to[len("mapping-") :]
    entries_by_id = {
        entry.id: entry for entry in db.scalars(select(MappingEntry)).all()
    }

    added = updated = 0

    for key in list(form.keys()):
        if not key.startswith("target_"):
            continue
        suffix = key[len("target_") :]
        if not suffix.isdigit():
            continue
        entry_id = int(suffix)
        entry = entries_by_id.get(entry_id)
        if entry is None:
            continue
        target = str(form.get(key, "")).strip()
        if target and target != entry.target:
            entry.target = target
            updated += 1

    seen: set[tuple[str, str, int, str]] = set()
    for key in list(form.keys()):
        if not key.startswith("new_target_"):
            continue
        suffix = key[len("new_target_") :]
        target = str(form.get(key, "")).strip()
        if not target:
            continue
        source_text = str(form.get(f"new_key_{suffix}", "")).strip()
        kind = str(form.get(f"new_kind_{suffix}", "")).strip()
        account_code = str(
            form.get(f"new_account_{suffix}", "")
        ).strip()
        year_text = str(form.get(f"new_year_{suffix}", "")).strip()
        year = int(year_text) if year_text.isdigit() else 0
        if kind not in ALLOWED_MAPPING_KINDS or not source_text:
            continue
        source_key = normalize_key(source_text)
        if not source_key or (kind, account_code, year, source_key) in seen:
            continue
        seen.add((kind, account_code, year, source_key))
        entry = db.scalar(
            select(MappingEntry).where(
                MappingEntry.kind == kind,
                MappingEntry.account_code == account_code,
                MappingEntry.year == year,
                MappingEntry.source_key == source_key,
            )
        )
        if entry is None:
            db.add(
                MappingEntry(
                    kind=kind,
                    account_code=account_code,
                    year=year,
                    source_key=source_key,
                    source_text=source_text,
                    target=target,
                )
            )
            added += 1
        elif target != entry.target:
            entry.target = target
            entry.source_text = source_text
            updated += 1

    if not (added or updated):
        return _redirect(redirect_path, view=view, account=account, kind=return_kind, status=return_status, q=return_q, page=return_page, _fragment=anchor, ok="Изменений нет.")

    try:
        db.commit()
    except Exception:  # noqa: BLE001 - show a friendly error
        db.rollback()
        log.exception("Failed to save mapping table")
        return _redirect(
            redirect_path,
            view=view,
            account=account,
            kind=return_kind,
            status=return_status,
            q=return_q,
            page=return_page,
            _fragment=anchor,
            err="Не удалось сохранить изменения мэппинга.",
        )

    try:
        rebuild_mapped_rows(db)
    except Exception:  # noqa: BLE001
        db.rollback()
        log.exception("Failed to rebuild mapping after table save")
        return _redirect(
            redirect_path,
            view=view,
            account=account,
            kind=return_kind,
            status=return_status,
            q=return_q,
            page=return_page,
            _fragment=anchor,
            err="Мэппинг сохранён, но пересчёт строк не удался.",
        )

    return _redirect(
        redirect_path,
        view=view,
        account=account,
        kind=return_kind,
        status=return_status,
        q=return_q,
        page=return_page,
        _fragment=anchor,
        ok=f"Сохранено: добавлено {added}, обновлено {updated}.",
    )


@app.get("/mapping/template.xlsx")
def mapping_template():
    data = build_mapping_template()
    return Response(
        content=data,
        media_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        headers={
            "Content-Disposition": 'attachment; filename="mapping_template.xlsx"'
        },
    )


@app.get("/mapping/export.xlsx")
def mapping_export(db: Session = Depends(get_db)):
    entries = db.scalars(
        select(MappingEntry).order_by(
            MappingEntry.account_code,
            MappingEntry.kind,
            MappingEntry.source_text,
        )
    ).all()
    data = build_mapping_export(entries)
    return Response(
        content=data,
        media_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        headers={
            "Content-Disposition": 'attachment; filename="mapping_current.xlsx"'
        },
    )


@app.get("/mapping/fot/template.xlsx")
def mapping_fot_template():
    data = build_fot_mapping_template()
    return Response(
        content=data,
        media_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        headers={
            "Content-Disposition": 'attachment; filename="mapping_fot_template.xlsx"'
        },
    )


@app.get("/mapping/fot/export.xlsx")
def mapping_fot_export(db: Session = Depends(get_db)):
    uploads = db.scalars(
        select(Upload)
        .where(Upload.account_code.in_(("26", "90.02")))
        .distinct(Upload.account_code)
        .order_by(Upload.account_code, Upload.created_at.desc())
    ).all()
    entries = db.scalars(
        select(MappingEntry)
        .where(MappingEntry.kind == "department")
        .order_by(MappingEntry.source_text)
    ).all()
    rows = []
    for upload in uploads:
        for row in _fot_mapping_rows(db, entries, upload):
            rows.append(
                {
                    "account": row.account,
                    "year": row.year,
                    "department": row.department,
                    "target": row.target,
                }
            )
    data = build_fot_mapping_export(rows)
    return Response(
        content=data,
        media_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        headers={
            "Content-Disposition": 'attachment; filename="mapping_fot_current.xlsx"'
        },
    )


@app.post("/mapping/import")
async def mapping_import(
    file: UploadFile = File(...),
    account: str = Form(""),
    view: str = Form("main"),
    db: Session = Depends(get_db),
):
    view = view if view in ("main", "fot") else "main"
    filename = Path(file.filename or "").name
    extension = Path(filename).suffix.lower()
    if extension not in ALLOWED_EXTENSIONS:
        return _redirect(
            "/mapping",
            view=view,
            account=account,
            err="Для импорта нужен файл .xlsx или .xlsm.",
        )
    data = await file.read()
    if not data:
        return _redirect("/mapping", view=view, account=account, err="Файл пустой.")
    if len(data) > MAX_UPLOAD_BYTES:
        return _redirect(
            "/mapping",
            view=view,
            account=account,
            err=f"Файл больше {MAX_UPLOAD_BYTES // (1024 * 1024)} МБ.",
        )

    try:
        items, _skipped, conflicts = parse_mapping_workbook(data)
    except MappingImportError as exc:
        return _redirect("/mapping", view=view, account=account, err=str(exc))

    try:
        added, updated = apply_mapping_entries(db, items)
    except Exception:  # noqa: BLE001 - show a friendly error
        db.rollback()
        log.exception("Failed to import mapping from %s", filename)
        return _redirect(
            "/mapping",
            view=view,
            account=account,
            err="Не удалось сохранить импортированный мэппинг.",
        )

    try:
        rebuild_mapped_rows(db)
    except Exception:  # noqa: BLE001
        db.rollback()
        log.exception("Failed to rebuild mapping after import")
        return _redirect(
            "/mapping",
            view=view,
            account=account,
            err="Мэппинг сохранён, но пересчёт строк не удался.",
        )

    message = f"Импорт: добавлено {added}, обновлено {updated}."
    if conflicts:
        message += (
            f" Конфликтов (разные значения для одного ключа): {conflicts}"
            " — использовано последнее."
        )
    message += " Мэппинг пересчитан."
    return _redirect(
        "/mapping",
        view=view,
        account=account,
        ok=message,
        _fragment="fot" if view == "fot" else "unified",
    )


@app.post("/mapping/add")
def mapping_add(
    kind: str = Form(...),
    source_text: str = Form(...),
    target: str = Form(...),
    account: str = Form(""),
    db: Session = Depends(get_db),
):
    source_text = source_text.strip()
    target = target.strip()
    if kind not in ALLOWED_MAPPING_KINDS:
        return _redirect(
            "/mapping", account=account, err="Неизвестный тип мэппинга."
        )
    if not source_text or not target:
        return _redirect(
            "/mapping",
            account=account,
            err="Заполните значение и мэппинг.",
        )
    key = normalize_key(source_text)
    if not key:
        return _redirect(
            "/mapping", account=account, err="Пустое значение мэппинга."
        )

    entry = db.scalar(
        select(MappingEntry).where(
            MappingEntry.kind == kind, MappingEntry.source_key == key
        )
    )
    if entry is None:
        db.add(
            MappingEntry(
                kind=kind,
                source_key=key,
                source_text=source_text,
                target=target,
            )
        )
    else:
        entry.source_text = source_text
        entry.target = target
    db.commit()

    try:
        rebuild_mapped_rows(db)
    except Exception:  # noqa: BLE001 - mapping is saved, rows can be rebuilt
        db.rollback()
        log.exception("Failed to rebuild mapping")
        return _redirect(
            "/mapping",
            account=account,
            err="Мэппинг сохранён, но пересчёт строк не удался.",
        )
    return _redirect(
        "/mapping",
        account=account,
        ok="Мэппинг сохранён.",
    )


@app.post("/mapping/rebuild")
def mapping_rebuild(
    account: str = Form(""),
    db: Session = Depends(get_db),
):
    try:
        rebuild_mapped_rows(db)
    except Exception:  # noqa: BLE001 - show a friendly error
        db.rollback()
        log.exception("Failed to rebuild mapping")
        return _redirect(
            "/mapping", account=account, err="Не удалось пересчитать мэппинг."
        )
    return _redirect(
        "/mapping",
        account=account,
        ok="Мэппинг пересчитан.",
    )


@app.post("/mapping/{entry_id}/update")
def mapping_update(
    entry_id: int,
    target: str = Form(...),
    account: str = Form(""),
    db: Session = Depends(get_db),
):
    entry = db.get(MappingEntry, entry_id)
    if entry is None:
        raise HTTPException(status_code=404, detail="Мэппинг не найден")
    target = target.strip()
    if not target:
        return _redirect(
            "/mapping", account=account, err="Мэппинг не может быть пустым."
        )
    entry.target = target
    db.commit()
    try:
        rebuild_mapped_rows(db)
    except Exception:  # noqa: BLE001
        db.rollback()
        log.exception("Failed to rebuild mapping")
        return _redirect(
            "/mapping",
            account=account,
            err="Мэппинг сохранён, но пересчёт строк не удался.",
        )
    return _redirect(
        "/mapping",
        account=account,
        ok="Мэппинг обновлён.",
    )


UPLOAD_STATUS_LABELS = {
    "ok": "Загружено",
    "duplicate": "Дубликат",
    "error": "Отклонено",
}


def _upload_page_data(db: Session) -> dict:
    """Latest upload per account, full history and headline counters."""
    accounts = db.scalars(
        select(Account).order_by(Account.sort_order, Account.code)
    ).all()
    latest: list[Upload] = []
    for account in accounts:
        upload = _latest_upload(db, account.code)
        if upload is not None:
            latest.append(upload)
    latest.sort(
        key=lambda item: item.created_at.timestamp() if item.created_at else 0,
        reverse=True,
    )
    history = db.scalars(
        select(Upload).order_by(Upload.created_at.desc()).limit(40)
    ).all()
    total_files = db.scalar(select(func.count(Upload.id))) or 0
    total_rows = db.scalar(select(func.coalesce(func.sum(Upload.row_count), 0))) or 0
    return {
        "recent": latest,
        "history": history,
        "summary": {
            "total_files": total_files,
            "accounts": len(latest),
            "rows": total_rows,
            "last": history[0].created_at if history else None,
        },
    }


@app.get("/upload", response_class=HTMLResponse)
def upload_page(
    request: Request,
    ok: str | None = None,
    err: str | None = None,
    db: Session = Depends(get_db),
):
    context = {
        "active_nav": "upload",
        "max_mb": MAX_UPLOAD_BYTES // (1024 * 1024),
        "ok": ok,
        "err": err,
        "results": None,
        "batch": None,
        **_upload_page_data(db),
    }
    return templates.TemplateResponse(request, "upload.html", context)


async def _process_upload(db: Session, file: UploadFile) -> dict:
    """Parse and store one uploaded workbook, returning a table row."""
    filename = Path(file.filename or "").name
    result: dict = {
        "filename": filename or "Без имени",
        "status": "error",
        "label": UPLOAD_STATUS_LABELS["error"],
        "account": None,
        "period": None,
        "company": None,
        "rows": 0,
        "debit": None,
        "credit": None,
        "upload_id": None,
        "message": "",
        "existing_at": None,
    }
    if not filename:
        result["message"] = "Файл не выбран."
        return result
    extension = Path(filename).suffix.lower()
    if extension not in ALLOWED_EXTENSIONS:
        result["message"] = (
            f"Неподдерживаемый формат «{extension or 'без расширения'}». "
            "Нужен файл .xlsx или .xlsm."
        )
        return result

    data = await file.read()
    if not data:
        result["message"] = "Файл пустой."
        return result
    if len(data) > MAX_UPLOAD_BYTES:
        result["message"] = f"Файл больше {MAX_UPLOAD_BYTES // (1024 * 1024)} МБ."
        return result

    try:
        report = parse_osv(data, filename)
    except ParseError as exc:
        result["message"] = str(exc)
        return result

    result.update(
        account=report.account_code,
        period=report.period,
        company=report.company,
        rows=len(report.rows),
        debit=report.totals[2],
        credit=report.totals[3],
    )

    digest = hashlib.sha256(data).hexdigest()
    existing = db.scalars(
        select(Upload).where(
            Upload.account_code == report.account_code,
            Upload.file_sha256 == digest,
        )
    ).first()
    if existing is not None:
        result.update(
            status="duplicate",
            label=UPLOAD_STATUS_LABELS["duplicate"],
            upload_id=existing.id,
            existing_at=existing.created_at,
        )
        return result

    upload = Upload(
        account_code=report.account_code,
        filename=filename,
        stored_name="pending",
        file_sha256=digest,
        company=report.company,
        period=report.period,
        source=report.source,
        row_count=len(report.rows),
        total_opening_debit=report.totals[0],
        total_opening_credit=report.totals[1],
        total_turnover_debit=report.totals[2],
        total_turnover_credit=report.totals[3],
        total_closing_debit=report.totals[4],
        total_closing_credit=report.totals[5],
    )

    stored_path: Path | None = None
    try:
        db.add(upload)
        db.flush()

        stored_path = UPLOAD_DIR / f"{upload.id}{extension}"
        stored_path.write_bytes(data)
        upload.stored_name = stored_path.name

        objects: list[OsvRow] = []
        for parsed in report.rows:
            row = OsvRow(
                upload_id=upload.id,
                position=parsed.position,
                level=parsed.level,
                node_type=parsed.node_type,
                code=parsed.code,
                name=parsed.name,
                month=parsed.month,
                opening_debit=parsed.amounts[0],
                opening_credit=parsed.amounts[1],
                turnover_debit=parsed.amounts[2],
                turnover_credit=parsed.amounts[3],
                closing_debit=parsed.amounts[4],
                closing_credit=parsed.amounts[5],
            )
            if parsed.parent_index is not None:
                row.parent = objects[parsed.parent_index]
            db.add(row)
            objects.append(row)

        db.commit()
        try:
            mapped = rebuild_mapped_rows(db, upload_id=upload.id)
            log.info("Mapping built for upload %s: %s rows", upload.id, mapped)
        except Exception:  # noqa: BLE001 - upload itself is already committed
            db.rollback()
            log.exception("Failed to build mapping for upload %s", upload.id)
    except Exception as exc:  # noqa: BLE001 - surface a friendly error
        db.rollback()
        if stored_path is not None:
            stored_path.unlink(missing_ok=True)
        log.exception("Failed to import %s", filename)
        result["message"] = f"Не удалось сохранить данные в базу: {exc}"
        return result

    log.info(
        "Imported %s for account %s (%s rows)",
        filename,
        report.account_code,
        len(report.rows),
    )
    result.update(
        status="ok",
        label=UPLOAD_STATUS_LABELS["ok"],
        upload_id=upload.id,
    )
    return result


@app.post("/upload", response_class=HTMLResponse)
async def upload_file(
    request: Request,
    files: list[UploadFile] | None = File(default=None),
    file: UploadFile | None = File(default=None),
    db: Session = Depends(get_db),
):
    incoming = [item for item in (files or []) if item.filename]
    if file is not None and file.filename:
        incoming.append(file)

    context = {
        "active_nav": "upload",
        "max_mb": MAX_UPLOAD_BYTES // (1024 * 1024),
        "ok": None,
        "err": None,
        "results": None,
        "batch": None,
        **_upload_page_data(db),
    }
    if not incoming:
        context["err"] = "Файл не выбран."
        return templates.TemplateResponse(request, "upload.html", context)

    results = []
    for item in incoming:
        results.append(await _process_upload(db, item))
    context.update(
        results=results,
        batch={
            "total": len(results),
            "ok": sum(1 for item in results if item["status"] == "ok"),
            "error": sum(1 for item in results if item["status"] == "error"),
            "duplicate": sum(1 for item in results if item["status"] == "duplicate"),
        },
        **_upload_page_data(db),
    )
    return templates.TemplateResponse(request, "upload.html", context)


@app.post("/uploads/{upload_id}/delete")
def delete_upload(upload_id: str, db: Session = Depends(get_db)):
    upload = db.get(Upload, upload_id)
    if upload is None:
        raise HTTPException(status_code=404, detail="Загрузка не найдена")
    account_code = upload.account_code
    stored_path = UPLOAD_DIR / upload.stored_name
    filename = upload.filename
    db.delete(upload)
    db.commit()
    stored_path.unlink(missing_ok=True)
    return _redirect(
        f"/accounts/{account_code}",
        ok=f"Загрузка «{filename}» и её строки удалены.",
    )


@app.get("/uploads/{upload_id}/file")
def download_upload(upload_id: str, db: Session = Depends(get_db)):
    upload = db.get(Upload, upload_id)
    if upload is None:
        raise HTTPException(status_code=404, detail="Загрузка не найдена")
    path = UPLOAD_DIR / upload.stored_name
    if not path.exists():
        raise HTTPException(status_code=404, detail="Файл не найден на диске")
    return FileResponse(
        path,
        filename=upload.filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/api/accounts")
def api_accounts(db: Session = Depends(get_db)):
    accounts = db.scalars(
        select(Account).order_by(Account.sort_order, Account.code)
    ).all()
    result = []
    for account in accounts:
        upload = _latest_upload(db, account.code)
        result.append(
            {
                "code": account.code,
                "name": account.name,
                "is_ready": account.is_ready,
                "uploads": db.scalar(
                    select(func.count(Upload.id)).where(
                        Upload.account_code == account.code
                    )
                )
                or 0,
                "latest_upload": (
                    {
                        "id": upload.id,
                        "filename": upload.filename,
                        "period": upload.period,
                        "created_at": upload.created_at.isoformat(),
                        "row_count": upload.row_count,
                        "turnover_debit": float(upload.total_turnover_debit),
                        "turnover_credit": float(upload.total_turnover_credit),
                    }
                    if upload is not None
                    else None
                ),
            }
        )
    return {"accounts": result}


@app.get("/api/accounts/{code}")
def api_account(
    code: str, upload_id: str | None = None, db: Session = Depends(get_db)
):
    account = db.get(Account, code)
    if account is None:
        raise HTTPException(status_code=404, detail="Счёт не найден")

    uploads = db.scalars(
        select(Upload)
        .where(Upload.account_code == code)
        .order_by(Upload.created_at.desc())
    ).all()
    upload = next((item for item in uploads if item.id == upload_id), None)
    if upload is None and uploads:
        upload = uploads[0]

    rows = []
    if upload is not None:
        db_rows = db.scalars(
            select(OsvRow)
            .where(OsvRow.upload_id == upload.id)
            .order_by(OsvRow.position)
        ).all()
        rows = [
            {
                "position": row.position,
                "level": row.level,
                "parent_id": row.parent_id,
                "node_type": row.node_type,
                "code": row.code,
                "name": row.name,
                "display": _display_name(row),
                "month": row.month,
                "opening_debit": float(row.opening_debit),
                "opening_credit": float(row.opening_credit),
                "turnover_debit": float(row.turnover_debit),
                "turnover_credit": float(row.turnover_credit),
                "closing_debit": float(row.closing_debit),
                "closing_credit": float(row.closing_credit),
            }
            for row in db_rows
        ]

    return {
        "account": {"code": account.code, "name": account.name},
        "uploads": [
            {
                "id": item.id,
                "filename": item.filename,
                "period": item.period,
                "created_at": item.created_at.isoformat(),
                "row_count": item.row_count,
            }
            for item in uploads
        ],
        "upload": (
            {
                "id": upload.id,
                "filename": upload.filename,
                "period": upload.period,
                "company": upload.company,
                "source": upload.source,
                "created_at": upload.created_at.isoformat(),
                "row_count": upload.row_count,
                "totals": {
                    "opening_debit": float(upload.total_opening_debit),
                    "opening_credit": float(upload.total_opening_credit),
                    "turnover_debit": float(upload.total_turnover_debit),
                    "turnover_credit": float(upload.total_turnover_credit),
                    "closing_debit": float(upload.total_closing_debit),
                    "closing_credit": float(upload.total_closing_credit),
                },
            }
            if upload is not None
            else None
        ),
        "rows": rows,
    }


@app.get("/healthz")
def healthz(db: Session = Depends(get_db)):
    db.execute(text("SELECT 1"))
    return JSONResponse({"status": "ok"})
