"""Synchronous, owner-isolated persistence for the Telegram bookkeeping assistant.

All public repository methods open their own SQLAlchemy session.  This makes the
repository safe to invoke through ``asyncio.to_thread`` from an async bot: no
session object is shared between requests or threads.
"""
from __future__ import annotations

from datetime import date, datetime, timezone
from decimal import Decimal, InvalidOperation
import json
from typing import Any

from sqlalchemy import (
    JSON,
    BigInteger,
    Boolean,
    CheckConstraint,
    Date,
    DateTime,
    ForeignKey,
    Integer,
    String,
    UniqueConstraint,
    create_engine,
    event,
    func,
    select,
)
from sqlalchemy.engine import Engine
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import DeclarativeBase, Mapped, Session, mapped_column, sessionmaker
from sqlalchemy.pool import StaticPool

from domain import DocumentData


SUPPORTED_CURRENCIES = frozenset({"RUB", "USD", "EUR", "KZT", "BYN", "UAH", "GEL", "GBP", "CNY"})
MAX_DOCUMENTS_PER_OWNER = 50
MAX_DOCUMENT_BYTES_PER_OWNER = 20 * 1024 * 1024
MAX_DOCUMENT_LIST_LIMIT = 100
MAX_ENTRIES_LIMIT = 10_000
MAX_MONEY_INTEGRAL_DIGITS = 12
MAX_SOURCE_KEY_LENGTH = 255


class Base(DeclarativeBase):
    """Base class for this module's ORM models."""


class ConsentRecord(Base):
    __tablename__ = "consents"

    owner_id: Mapped[str] = mapped_column(String(128), primary_key=True)
    enabled: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False)
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), nullable=False, default=lambda: datetime.now(timezone.utc)
    )


class StoredDocument(Base):
    __tablename__ = "documents"
    __table_args__ = (
        CheckConstraint("payload_size >= 0", name="ck_documents_payload_size_nonnegative"),
    )

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    owner_id: Mapped[str] = mapped_column(String(128), nullable=False, index=True)
    filename: Mapped[str] = mapped_column(String(512), nullable=False)
    title: Mapped[str] = mapped_column(String(1000), nullable=False)
    payload: Mapped[dict[str, Any]] = mapped_column(JSON, nullable=False)
    payload_size: Mapped[int] = mapped_column(BigInteger, nullable=False)
    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), nullable=False, default=lambda: datetime.now(timezone.utc), index=True
    )


class LedgerEntry(Base):
    __tablename__ = "ledger_entries"
    __table_args__ = (
        UniqueConstraint("owner_id", "source_key", name="uq_ledger_entries_owner_source_key"),
        CheckConstraint("kind IN ('income', 'expense')", name="ck_ledger_entries_kind"),
        CheckConstraint("amount_minor > 0", name="ck_ledger_entries_amount_minor_positive"),
    )

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    owner_id: Mapped[str] = mapped_column(String(128), nullable=False, index=True)
    kind: Mapped[str] = mapped_column(String(7), nullable=False)
    # Money is deliberately stored as an integer number of cents/kopecks.  This
    # avoids SQLite's NUMERIC affinity converting aggregate values to floats.
    amount_minor: Mapped[int] = mapped_column(BigInteger, nullable=False)
    currency: Mapped[str] = mapped_column(String(3), nullable=False)
    entry_date: Mapped[date] = mapped_column(Date, nullable=False, index=True)
    category: Mapped[str] = mapped_column(String(100), nullable=False)
    note: Mapped[str] = mapped_column(String(1000), nullable=False)
    source_key: Mapped[str] = mapped_column(String(MAX_SOURCE_KEY_LENGTH), nullable=False)
    voided: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False, index=True)
    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), nullable=False, default=lambda: datetime.now(timezone.utc)
    )
    voided_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)


class LedgerAudit(Base):
    """Append-only audit record for destructive-looking ledger actions."""

    __tablename__ = "ledger_audit"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    owner_id: Mapped[str] = mapped_column(String(128), nullable=False, index=True)
    entry_id: Mapped[int] = mapped_column(
        Integer, ForeignKey("ledger_entries.id", ondelete="RESTRICT"), nullable=False, index=True
    )
    action: Mapped[str] = mapped_column(String(32), nullable=False)
    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), nullable=False, default=lambda: datetime.now(timezone.utc)
    )


def _normalise_database_url(database_url: str) -> str:
    """Supply a usable SQLite default and select psycopg for PostgreSQL URLs."""
    value = (database_url or "").strip()
    if not value:
        return "sqlite:///telegram_accountant.db"
    if value == ":memory:":
        return "sqlite:///:memory:"
    if value.startswith("postgres://"):
        return "postgresql+psycopg://" + value[len("postgres://") :]
    if value.startswith("postgresql://"):
        return "postgresql+psycopg://" + value[len("postgresql://") :]
    return value


def _owner(owner_id: object) -> str:
    value = str(owner_id)
    if not value or len(value) > 128:
        raise ValueError("Некорректный идентификатор владельца")
    return value


def _checked_limit(limit: int, maximum: int, label: str) -> int:
    if isinstance(limit, bool) or not isinstance(limit, int) or not 1 <= limit <= maximum:
        raise ValueError(f"Лимит {label} должен быть от 1 до {maximum}")
    return limit


def _checked_date(value: date | None, label: str) -> date | None:
    if value is not None and (not isinstance(value, date) or isinstance(value, datetime)):
        raise ValueError(f"{label} должна быть датой")
    return value


def _as_utc(value: datetime) -> datetime:
    """SQLite returns naive values; expose all persisted timestamps as UTC."""
    return value.replace(tzinfo=timezone.utc) if value.tzinfo is None else value


def _minor_to_decimal(value: int) -> Decimal:
    return Decimal(value).scaleb(-2)


def _validate_amount(amount: Decimal) -> int:
    if isinstance(amount, bool) or isinstance(amount, float):
        raise ValueError("Сумма должна быть точным Decimal, а не float")
    try:
        decimal_amount = amount if isinstance(amount, Decimal) else Decimal(amount)
    except (InvalidOperation, TypeError, ValueError) as exc:
        raise ValueError("Некорректная сумма") from exc
    if not decimal_amount.is_finite() or decimal_amount <= 0:
        raise ValueError("Сумма должна быть положительной")
    exponent = decimal_amount.as_tuple().exponent
    if exponent < -2:
        raise ValueError("У суммы допускается не более двух знаков после запятой")
    # adjusted() is the zero-based count of digits before the decimal point.
    if decimal_amount.adjusted() + 1 > MAX_MONEY_INTEGRAL_DIGITS:
        raise ValueError("Сумма содержит более 12 цифр до запятой")
    try:
        minor = int((decimal_amount * 100).to_integral_exact())
    except (InvalidOperation, ValueError) as exc:
        raise ValueError("Некорректная сумма") from exc
    if minor <= 0:
        raise ValueError("Сумма должна быть положительной")
    return minor


def _validate_entry_values(
    kind: str,
    amount: Decimal,
    currency: str,
    entry_date: date,
    category: str,
    note: str,
    source_key: str,
) -> tuple[int, str]:
    if kind not in {"income", "expense"}:
        raise ValueError("Тип операции должен быть income или expense")
    if not isinstance(currency, str) or currency.upper() not in SUPPORTED_CURRENCIES:
        raise ValueError("Поддерживается только один из разрешённых кодов ISO 4217")
    if currency != currency.upper() or len(currency) != 3:
        raise ValueError("Код валюты должен быть трёхбуквенным ISO-кодом в верхнем регистре")
    _checked_date(entry_date, "Дата операции")
    if not isinstance(category, str) or len(category) > 100:
        raise ValueError("Категория не должна превышать 100 символов")
    if not isinstance(note, str) or len(note) > 1000:
        raise ValueError("Примечание не должно превышать 1000 символов")
    if not isinstance(source_key, str) or not source_key or len(source_key) > MAX_SOURCE_KEY_LENGTH:
        raise ValueError(f"source_key должен содержать от 1 до {MAX_SOURCE_KEY_LENGTH} символов")
    return _validate_amount(amount), currency


def _document_payload(data: DocumentData) -> tuple[DocumentData, dict[str, Any], int]:
    parsed = DocumentData.model_validate(data)
    payload = parsed.model_dump(mode="json")
    encoded = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    return parsed, payload, len(encoded)


class Repository:
    """SQLAlchemy 2 synchronous repository with strict owner isolation.

    ``database_url`` accepts SQLite URLs and PostgreSQL URLs.  An empty value
    selects ``sqlite:///telegram_accountant.db``.  PostgreSQL URLs are routed to
    SQLAlchemy's ``psycopg`` dialect automatically.
    """

    def __init__(self, database_url: str):
        self.database_url = _normalise_database_url(database_url)
        engine_options: dict[str, Any] = {"future": True, "pool_pre_ping": True}
        if self.database_url.startswith("sqlite"):
            engine_options["connect_args"] = {"check_same_thread": False}
            if self.database_url in {"sqlite://", "sqlite:///:memory:"}:
                # One shared in-memory connection keeps a test database alive
                # across the per-call sessions.
                engine_options["poolclass"] = StaticPool
        self.engine: Engine = create_engine(self.database_url, **engine_options)
        if self.engine.dialect.name == "sqlite":
            self._configure_sqlite(self.engine)
        self._sessions = sessionmaker(bind=self.engine, class_=Session, expire_on_commit=False, future=True)

    @staticmethod
    def _configure_sqlite(engine: Engine) -> None:
        @event.listens_for(engine, "connect")
        def _sqlite_pragmas(dbapi_connection: Any, _connection_record: Any) -> None:
            cursor = dbapi_connection.cursor()
            try:
                # WAL and a nonzero timeout make short concurrent bot requests
                # substantially less likely to fail with "database is locked".
                cursor.execute("PRAGMA journal_mode=WAL")
                cursor.execute("PRAGMA busy_timeout=5000")
                cursor.execute("PRAGMA foreign_keys=ON")
            finally:
                cursor.close()

    def initialize(self) -> None:
        Base.metadata.create_all(self.engine)

    def close(self) -> None:
        self.engine.dispose()

    def consent(self, owner_id: object) -> bool:
        owner = _owner(owner_id)
        with self._sessions() as session:
            record = session.scalar(select(ConsentRecord).where(ConsentRecord.owner_id == owner))
            return bool(record and record.enabled)

    def set_consent(self, owner_id: object, enabled: bool) -> None:
        owner = _owner(owner_id)
        if not isinstance(enabled, bool):
            raise ValueError("Значение согласия должно быть bool")
        with self._sessions() as session:
            with session.begin():
                record = session.scalar(select(ConsentRecord).where(ConsentRecord.owner_id == owner))
                if record is None:
                    session.add(ConsentRecord(owner_id=owner, enabled=enabled))
                else:
                    record.enabled = enabled
                    record.updated_at = datetime.now(timezone.utc)

    def add_document(self, owner_id: object, filename: str, data: DocumentData) -> int:
        owner = _owner(owner_id)
        if not isinstance(filename, str) or not filename or len(filename) > 512:
            raise ValueError("Имя файла должно содержать от 1 до 512 символов")
        parsed, payload, payload_size = _document_payload(data)
        if payload_size > MAX_DOCUMENT_BYTES_PER_OWNER:
            raise ValueError("Документ превышает лимит 20 МБ")
        with self._sessions() as session:
            with session.begin():
                document_count, total_size = session.execute(
                    select(
                        func.count(StoredDocument.id),
                        func.coalesce(func.sum(StoredDocument.payload_size), 0),
                    ).where(StoredDocument.owner_id == owner)
                ).one()
                if int(document_count) >= MAX_DOCUMENTS_PER_OWNER:
                    raise ValueError(f"Для одного пользователя допускается не более {MAX_DOCUMENTS_PER_OWNER} документов")
                if int(total_size) + payload_size > MAX_DOCUMENT_BYTES_PER_OWNER:
                    raise ValueError("Превышен общий лимит хранения документов 20 МБ")
                document = StoredDocument(
                    owner_id=owner,
                    filename=filename,
                    title=parsed.title,
                    payload=payload,
                    payload_size=payload_size,
                )
                session.add(document)
                session.flush()
                return int(document.id)

    def get_document(self, owner_id: object, document_id: int) -> dict[str, Any] | None:
        owner = _owner(owner_id)
        with self._sessions() as session:
            document = session.scalar(
                select(StoredDocument).where(StoredDocument.owner_id == owner, StoredDocument.id == document_id)
            )
            return self._document_detail(document) if document is not None else None

    def list_documents(self, owner_id: object, limit: int = 30) -> list[dict[str, Any]]:
        owner = _owner(owner_id)
        checked_limit = _checked_limit(limit, MAX_DOCUMENT_LIST_LIMIT, "документов")
        with self._sessions() as session:
            documents = session.scalars(
                select(StoredDocument)
                .where(StoredDocument.owner_id == owner)
                .order_by(StoredDocument.created_at.desc(), StoredDocument.id.desc())
                .limit(checked_limit)
            ).all()
            return [
                {
                    "id": document.id,
                    "filename": document.filename,
                    "title": document.title,
                    "created_at": _as_utc(document.created_at),
                }
                for document in documents
            ]

    def documents_for_search(
        self, owner_id: object, document_id: int | None = None, limit: int = 50
    ) -> list[dict[str, Any]]:
        owner = _owner(owner_id)
        checked_limit = _checked_limit(limit, MAX_DOCUMENT_LIST_LIMIT, "документов")
        with self._sessions() as session:
            statement = select(StoredDocument).where(StoredDocument.owner_id == owner)
            if document_id is not None:
                statement = statement.where(StoredDocument.id == document_id)
            documents = session.scalars(
                statement.order_by(StoredDocument.created_at.desc(), StoredDocument.id.desc()).limit(checked_limit)
            ).all()
            return [self._document_detail(document) for document in documents]

    def delete_document(self, owner_id: object, document_id: int) -> bool:
        owner = _owner(owner_id)
        with self._sessions() as session:
            with session.begin():
                document = session.scalar(
                    select(StoredDocument).where(StoredDocument.owner_id == owner, StoredDocument.id == document_id)
                )
                if document is None:
                    return False
                session.delete(document)
                return True

    @staticmethod
    def _document_detail(document: StoredDocument) -> dict[str, Any]:
        return {
            "id": document.id,
            "filename": document.filename,
            "data": DocumentData.model_validate(document.payload),
            "created_at": _as_utc(document.created_at),
        }

    def add_entry(
        self,
        owner_id: object,
        kind: str,
        amount: Decimal,
        currency: str,
        entry_date: date,
        category: str,
        note: str,
        source_key: str,
    ) -> tuple[int, bool]:
        owner = _owner(owner_id)
        amount_minor, currency = _validate_entry_values(
            kind, amount, currency, entry_date, category, note, source_key
        )
        with self._sessions() as session:
            # Fast path preserves idempotence without attempting an insert.
            existing = session.scalar(
                select(LedgerEntry).where(LedgerEntry.owner_id == owner, LedgerEntry.source_key == source_key)
            )
            if existing is not None:
                return int(existing.id), False
            entry = LedgerEntry(
                owner_id=owner,
                kind=kind,
                amount_minor=amount_minor,
                currency=currency,
                entry_date=entry_date,
                category=category,
                note=note,
                source_key=source_key,
            )
            try:
                # The preceding SELECT may have opened SQLAlchemy's implicit
                # transaction, so commit that transaction instead of opening a
                # nested ``session.begin()`` block.
                session.add(entry)
                session.commit()
                entry_id = int(entry.id)
                return entry_id, True
            except IntegrityError:
                # A concurrent call may have inserted the same owner/source key
                # after the fast-path read.  Query it only in that owner's scope.
                session.rollback()
                existing = session.scalar(
                    select(LedgerEntry).where(LedgerEntry.owner_id == owner, LedgerEntry.source_key == source_key)
                )
                if existing is not None:
                    return int(existing.id), False
                raise

    def entries(
        self,
        owner_id: object,
        start: date | None = None,
        end: date | None = None,
        limit: int = MAX_ENTRIES_LIMIT,
    ) -> list[dict[str, Any]]:
        owner = _owner(owner_id)
        start = _checked_date(start, "Начальная дата")
        end = _checked_date(end, "Конечная дата")
        if start is not None and end is not None and start > end:
            raise ValueError("Начальная дата не может быть позже конечной")
        checked_limit = _checked_limit(limit, MAX_ENTRIES_LIMIT, "операций")
        with self._sessions() as session:
            statement = select(LedgerEntry).where(LedgerEntry.owner_id == owner)
            if start is not None:
                statement = statement.where(LedgerEntry.entry_date >= start)
            if end is not None:
                statement = statement.where(LedgerEntry.entry_date <= end)
            result = session.scalars(
                statement.order_by(LedgerEntry.entry_date.desc(), LedgerEntry.id.desc()).limit(checked_limit + 1)
            ).all()
            if len(result) > checked_limit:
                raise ValueError(f"Журнал содержит более {checked_limit} операций; сузьте период для полной выгрузки")
            return [self._entry_dict(entry) for entry in result]

    def void_entry(self, owner_id: object, entry_id: int) -> bool:
        owner = _owner(owner_id)
        with self._sessions() as session:
            with session.begin():
                entry = session.scalar(
                    select(LedgerEntry).where(LedgerEntry.owner_id == owner, LedgerEntry.id == entry_id)
                )
                if entry is None or entry.voided:
                    return False
                entry.voided = True
                entry.voided_at = datetime.now(timezone.utc)
                session.add(LedgerAudit(owner_id=owner, entry_id=entry.id, action="void"))
                return True

    def summary(
        self, owner_id: object, start: date | None = None, end: date | None = None
    ) -> dict[str, dict[str, Decimal | int]]:
        """Return exact, per-currency totals without SQL NUMERIC aggregation."""
        owner = _owner(owner_id)
        start = _checked_date(start, "Начальная дата")
        end = _checked_date(end, "Конечная дата")
        if start is not None and end is not None and start > end:
            raise ValueError("Начальная дата не может быть позже конечной")
        with self._sessions() as session:
            statement = (
                select(LedgerEntry.currency, LedgerEntry.kind, LedgerEntry.amount_minor)
                .where(LedgerEntry.owner_id == owner, LedgerEntry.voided.is_(False))
                .order_by(LedgerEntry.id)
            )
            if start is not None:
                statement = statement.where(LedgerEntry.entry_date >= start)
            if end is not None:
                statement = statement.where(LedgerEntry.entry_date <= end)
            minor_totals: dict[str, dict[str, int]] = {}
            # Iterate database integer amounts and sum in Python.  This is
            # intentional: SQLite SUM over NUMERIC values may coerce to float.
            for currency, kind, amount_minor in session.execute(statement):
                bucket = minor_totals.setdefault(currency, {"income": 0, "expense": 0, "count": 0})
                bucket[kind] += int(amount_minor)
                bucket["count"] += 1
        return {
            currency: {
                "income": _minor_to_decimal(values["income"]),
                "expense": _minor_to_decimal(values["expense"]),
                "balance": _minor_to_decimal(values["income"] - values["expense"]),
                "count": values["count"],
            }
            for currency, values in sorted(minor_totals.items())
        }

    @staticmethod
    def _entry_dict(entry: LedgerEntry) -> dict[str, Any]:
        return {
            "id": entry.id,
            "kind": entry.kind,
            "amount": _minor_to_decimal(entry.amount_minor),
            "currency": entry.currency,
            "entry_date": entry.entry_date,
            "category": entry.category,
            "note": entry.note,
            "voided": entry.voided,
        }
