from __future__ import annotations

from datetime import datetime
from decimal import Decimal
from uuid import uuid4

from sqlalchemy import (
    BigInteger,
    Boolean,
    DateTime,
    ForeignKey,
    Index,
    Integer,
    JSON,
    Numeric,
    String,
    Text,
    UniqueConstraint,
    func,
)
from sqlalchemy.orm import Mapped, mapped_column, relationship

from .db import Base

MONEY = Numeric(20, 2)


class Account(Base):
    __tablename__ = "accounts"

    code: Mapped[str] = mapped_column(String(10), primary_key=True)
    name: Mapped[str] = mapped_column(String(200))
    description: Mapped[str | None] = mapped_column(Text)
    sort_order: Mapped[int] = mapped_column(Integer, default=0)
    is_ready: Mapped[bool] = mapped_column(Boolean, default=False)

    uploads: Mapped[list[Upload]] = relationship(
        back_populates="account",
        cascade="all, delete-orphan",
        order_by="Upload.created_at.desc()",
    )


class Upload(Base):
    __tablename__ = "uploads"
    __table_args__ = (
        UniqueConstraint("account_code", "file_sha256", name="uq_upload_account_hash"),
    )

    id: Mapped[str] = mapped_column(
        String(36), primary_key=True, default=lambda: str(uuid4())
    )
    account_code: Mapped[str] = mapped_column(
        ForeignKey("accounts.code", ondelete="CASCADE"), index=True
    )
    filename: Mapped[str] = mapped_column(String(255))
    stored_name: Mapped[str] = mapped_column(String(80))
    file_sha256: Mapped[str] = mapped_column(String(64), index=True)
    company: Mapped[str | None] = mapped_column(String(255))
    period: Mapped[str | None] = mapped_column(String(120))
    source: Mapped[str | None] = mapped_column(String(255))
    row_count: Mapped[int] = mapped_column(Integer, default=0)

    total_opening_debit: Mapped[Decimal] = mapped_column(MONEY, default=0)
    total_opening_credit: Mapped[Decimal] = mapped_column(MONEY, default=0)
    total_turnover_debit: Mapped[Decimal] = mapped_column(MONEY, default=0)
    total_turnover_credit: Mapped[Decimal] = mapped_column(MONEY, default=0)
    total_closing_debit: Mapped[Decimal] = mapped_column(MONEY, default=0)
    total_closing_credit: Mapped[Decimal] = mapped_column(MONEY, default=0)

    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), server_default=func.now(), index=True
    )

    account: Mapped[Account] = relationship(back_populates="uploads")
    rows: Mapped[list[OsvRow]] = relationship(
        back_populates="upload",
        cascade="all, delete-orphan",
        order_by="OsvRow.position",
    )


class OsvRow(Base):
    __tablename__ = "osv_rows"

    id: Mapped[int] = mapped_column(BigInteger, primary_key=True, autoincrement=True)
    upload_id: Mapped[str] = mapped_column(
        ForeignKey("uploads.id", ondelete="CASCADE"), index=True
    )
    parent_id: Mapped[int | None] = mapped_column(
        ForeignKey("osv_rows.id", ondelete="CASCADE"), index=True
    )
    position: Mapped[int] = mapped_column(Integer)
    level: Mapped[int] = mapped_column(Integer)
    node_type: Mapped[str] = mapped_column(String(20))
    code: Mapped[str | None] = mapped_column(String(50))
    name: Mapped[str] = mapped_column(Text)
    month: Mapped[str | None] = mapped_column(String(40))

    opening_debit: Mapped[Decimal] = mapped_column(MONEY, default=0)
    opening_credit: Mapped[Decimal] = mapped_column(MONEY, default=0)
    turnover_debit: Mapped[Decimal] = mapped_column(MONEY, default=0)
    turnover_credit: Mapped[Decimal] = mapped_column(MONEY, default=0)
    closing_debit: Mapped[Decimal] = mapped_column(MONEY, default=0)
    closing_credit: Mapped[Decimal] = mapped_column(MONEY, default=0)

    upload: Mapped[Upload] = relationship(back_populates="rows")
    parent: Mapped[OsvRow | None] = relationship(
        remote_side="OsvRow.id", back_populates="children"
    )
    children: Mapped[list[OsvRow]] = relationship(
        back_populates="parent", cascade="all, delete-orphan"
    )


class MappingEntry(Base):
    __tablename__ = "mapping_entries"
    __table_args__ = (
        UniqueConstraint(
            "kind",
            "account_code",
            "year",
            "source_key",
            name="uq_mapping_kind_account_year_key",
        ),
    )

    id: Mapped[int] = mapped_column(BigInteger, primary_key=True, autoincrement=True)
    kind: Mapped[str] = mapped_column(String(20), index=True)
    account_code: Mapped[str] = mapped_column(String(20), default="", index=True)
    year: Mapped[int] = mapped_column(Integer, default=0, index=True)
    source_key: Mapped[str] = mapped_column(String(500))
    source_text: Mapped[str] = mapped_column(String(500))
    target: Mapped[str] = mapped_column(Text)
    payload: Mapped[dict | None] = mapped_column(JSON, nullable=True)
    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), server_default=func.now()
    )
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), server_default=func.now(), onupdate=func.now()
    )


class MappedRow(Base):
    __tablename__ = "mapped_rows"
    __table_args__ = (
        Index("ix_mapped_rows_account_position", "account_code", "position"),
    )

    id: Mapped[int] = mapped_column(BigInteger, primary_key=True, autoincrement=True)
    upload_id: Mapped[str] = mapped_column(
        ForeignKey("uploads.id", ondelete="CASCADE"), index=True
    )
    osv_row_id: Mapped[int] = mapped_column(
        ForeignKey("osv_rows.id", ondelete="CASCADE"), unique=True
    )
    account_code: Mapped[str] = mapped_column(
        ForeignKey("accounts.code", ondelete="CASCADE"), index=True
    )
    position: Mapped[int] = mapped_column(Integer)

    department: Mapped[str] = mapped_column(String(500), default="")
    group_name: Mapped[str] = mapped_column(String(500), default="")
    article: Mapped[str] = mapped_column(String(500), default="")
    month: Mapped[str | None] = mapped_column(String(40))

    debit: Mapped[Decimal] = mapped_column(MONEY, default=0)
    credit: Mapped[Decimal] = mapped_column(MONEY, default=0)

    target: Mapped[str | None] = mapped_column(Text)
    mapping_kind: Mapped[str | None] = mapped_column(String(20))
    mapping_key: Mapped[str | None] = mapped_column(String(500))
    status: Mapped[str] = mapped_column(String(20), index=True)

    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), server_default=func.now()
    )
