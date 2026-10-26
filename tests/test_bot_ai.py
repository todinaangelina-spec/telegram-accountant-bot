"""Offline tests: Telegram and OpenAI are never contacted."""
from datetime import date
from decimal import Decimal
from types import SimpleNamespace
from unittest.mock import AsyncMock

import pytest
from aiogram.types import Chat, Message, User

from ai_service import AIService, build_context
from bot import AccessMiddleware, Application, date_range, report_text
from config import Settings
from domain import DocumentData, TableData
from ingest import ParsedSource
from storage import Repository


@pytest.fixture
def settings():
    return Settings(bot_token="123456:TEST_ONLY_NOT_A_REAL_TOKEN", allowed_user_ids=frozenset({123}),
                    database_url="sqlite://", openai_api_key="")


@pytest.fixture
def repo(tmp_path):
    obj = Repository(f"sqlite:///{tmp_path / 'test.db'}")
    obj.initialize()
    yield obj
    obj.close()


def message(user=123, chat_type="private", text="/report"):
    return Message(message_id=1, date=0, chat=Chat(id=user, type=chat_type),
                   from_user=User(id=user, is_bot=False, first_name="Fixture"), text=text)


@pytest.mark.asyncio
async def test_access_denied(monkeypatch):
    reply = AsyncMock()
    monkeypatch.setattr(Message, "answer", reply)
    handler = AsyncMock()
    await AccessMiddleware(frozenset({123}))(handler, message(user=999), {})
    handler.assert_not_awaited()
    reply.assert_awaited_once()


@pytest.mark.asyncio
async def test_group_silently_rejected(monkeypatch):
    reply = AsyncMock()
    monkeypatch.setattr(Message, "answer", reply)
    handler = AsyncMock()
    await AccessMiddleware(frozenset({123}))(handler, message(chat_type="group"), {})
    handler.assert_not_awaited()
    reply.assert_not_awaited()


@pytest.mark.asyncio
async def test_whoami_accessible_without_financial_access(monkeypatch):
    reply = AsyncMock()
    monkeypatch.setattr(Message, "answer", reply)
    handler = AsyncMock()
    await AccessMiddleware(frozenset({123}))(handler, message(user=999, text="/whoami"), {})
    handler.assert_not_awaited()
    assert "999" in reply.call_args.args[0]


@pytest.mark.asyncio
async def test_app_requires_consent(repo, settings):
    ai = SimpleNamespace(available=True)
    app = Application(settings, repo, ai)
    with pytest.raises(ValueError, match="consent"):
        await app.require_ai(123)
    repo.set_consent(123, True)
    await app.require_ai(123)
    with pytest.raises(ValueError):
        await app.require_ai(456)


@pytest.mark.asyncio
async def test_get_document_cannot_cross_owner(repo, settings):
    data = DocumentData(title="Fixture", text="Fixture-only data", tables=[], warnings=[])
    ident = repo.add_document(456, "fixture.txt", data)
    app = Application(settings, repo, SimpleNamespace(available=False))
    with pytest.raises(ValueError, match="не найден"):
        await app.get_doc(123, str(ident))


def test_context_has_sources_and_scope():
    doc = {"id": 7, "filename": "fixture.csv", "data": DocumentData(
        title="Fixture", text="", warnings=[], tables=[TableData(
            name="Records", columns=["Категория", "Сумма"], rows=[["Аренда", "500.00"]])])}
    context, note = build_context("Аренда", [doc])
    assert "[D7:фрагмент1]" in context
    assert "500.00" in context
    assert "не полный анализ" in note


def test_retrieval_bounded_and_reports_partial():
    docs = [{"id": n, "filename": "fixture.txt", "data": DocumentData(
        title=f"Fixture{n}", text=("аренда fixture " * 700), tables=[], warnings=[])} for n in range(12)]
    context, note = build_context("аренда", docs, max_chars=4000)
    assert len(context) < 4100
    assert "из 12" in note


@pytest.mark.asyncio
async def test_ai_not_configured_is_explicit(settings):
    ai = AIService(settings)
    assert not ai.available
    source = ParsedSource(filename="fixture.txt", text="Fixture", tables=[], images=[], warnings=[])
    with pytest.raises(ValueError, match="не настроен"):
        await ai.extract(123, source)
    await ai.close()


@pytest.mark.asyncio
async def test_ai_extract_mock(settings):
    ai = AIService(settings)
    result = DocumentData(title="Fixture", text="fixture", tables=[], warnings=[])
    parse = AsyncMock(return_value=SimpleNamespace(choices=[SimpleNamespace(
        message=SimpleNamespace(refusal=None, parsed=result))]))
    ai.client = SimpleNamespace(chat=SimpleNamespace(completions=SimpleNamespace(parse=parse)))
    source = ParsedSource(filename="fixture.txt", text="fixture", tables=[], images=[], warnings=[])
    data = await ai.extract(123, source)
    assert "ИИ" in data.warnings[-1]
    assert parse.call_args.kwargs["store"] is False
    assert parse.call_args.kwargs["response_format"] is DocumentData


@pytest.mark.asyncio
async def test_ai_large_input_not_sent(settings):
    ai = AIService(settings)
    parse = AsyncMock()
    ai.client = SimpleNamespace(chat=SimpleNamespace(completions=SimpleNamespace(parse=parse)))
    source = ParsedSource(filename="fixture.txt", text="x" * 20001, tables=[], images=[], warnings=[])
    with pytest.raises(ValueError, match="разделите"):
        await ai.extract(123, source)
    parse.assert_not_awaited()


def test_date_range_and_currency_summary():
    assert date_range(None) == (None, None)
    assert date_range("2026-10-01 2026-10-31") == (date(2026, 10, 1), date(2026, 10, 31))
    with pytest.raises(ValueError):
        date_range("2026-10-31 2026-10-01")
    text = report_text({"USD": {"income": Decimal("0.30"), "expense": Decimal("0.10"),
                                "balance": Decimal("0.20"), "count": 3}})
    assert "0.20" in text or "0,20" in text
    assert "USD" in text
