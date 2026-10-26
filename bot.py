"""Telegram bookkeeping MVP. Run: python bot.py; no credentials in source."""
from __future__ import annotations

import asyncio
import logging
import os
import secrets
import tempfile
import time
from collections import defaultdict
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any, Awaitable, Callable
from zoneinfo import ZoneInfo

from aiogram import BaseMiddleware, Bot, Dispatcher, F, Router
from aiogram.filters import Command, CommandObject
from aiogram.types import (
    BotCommand,
    BufferedInputFile,
    KeyboardButton,
    Message,
    ReplyKeyboardMarkup,
    TelegramObject,
)

from accounting import money_text, parse_money, safe_calculate, vat
from ai_service import AIService
from config import Settings
from domain import DocumentData
from exports import chart_xlsx, document_xlsx, ledger_xlsx
from ingest import ParsedSource, parse_file, parse_text
from storage import Repository

LOGGER = logging.getLogger("accountant")
HELP = """Помощник по управленческому учёту

Файлы: PDF, DOCX, XLSX, CSV, TXT, JSON, JPG/PNG/WEBP.
/text исходные данные — сохранить текст и, при согласии, сделать таблицу.
/docs — последние 30 документов.
/use ID или /use all — выбрать файл или все свои файлы для вопросов.
/ask вопрос — ответить по выбранным документам и сводке журнала.
/table ID — сделать таблицу из сохранённого текста с помощью ИИ.
/export ID — выгрузить документ в Excel.
/chart ID ТАБЛИЦА СТОЛБЕЦ_КАТЕГОРИИ СТОЛБЕЦ_ЧИСЛА — Excel-график.
/sum ID номер_таблицы номер_столбца — точная сумма столбца (нумерация с 1).

Учёт (только ручные записи, автоматических проводок нет):
/income 1250.50 RUB 2026-10-25 Продажи | Оплата заказа
/expense 500 RUB 2026-10-25 Аренда | Офис
/report — сводка журнала, отдельно по каждой валюте.
/report 2026-10-01 2026-10-31 — сводка за период.
/ledger [начало конец] — журнал и итоги в Excel.
/void ID — аннулировать свою запись, не стирая историю.
/calc (1200 + 350) * 2 — точная арифметика.
/vat 1200 20 — выделить НДС из суммы и начислить сверху; ставка ваша.

/consent — разрешить внешнюю обработку ИИ.
/revoke — отключить следующие обращения к ИИ.
/privacy — сведения о хранении и передаче данных.
/forget ID — начать удаление документа с подтверждением.
/whoami — показать ваш Telegram ID.

Без согласия работают локальные таблицы, журнал и расчёты.
Фото/сканы требуют ИИ. Извлечённые числа надо сверять с оригиналом.
Это не полноценная регламентированная бухгалтерская система."""

MENU_UPLOAD = "Загрузить файл"
MENU_DOCS = "Мои документы"
MENU_TABLE = "Сделать таблицу / Excel"
MENU_CHART = "Построить график"
MENU_ASK = "Спросить ИИ"
MENU_REVIEW = "Проверить таблицу"
MENU_REPORT = "Финансовый отчёт"
MENU_PRIVACY = "Приватность"
MENU_HELP = "Помощь"
MAIN_MENU = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text=MENU_UPLOAD), KeyboardButton(text=MENU_DOCS)],
        [KeyboardButton(text=MENU_TABLE), KeyboardButton(text=MENU_CHART)],
        [KeyboardButton(text=MENU_ASK), KeyboardButton(text=MENU_REVIEW)],
        [KeyboardButton(text=MENU_REPORT), KeyboardButton(text=MENU_PRIVACY)],
        [KeyboardButton(text=MENU_HELP)],
    ],
    resize_keyboard=True,
    is_persistent=True,
    input_field_placeholder="Нажмите кнопку или задайте вопрос",
)


async def send_text(message: Message, text: str):
    # Plain text prevents untrusted file contents from becoming Telegram markup.
    for offset in range(0, len(text), 3500):
        await message.answer(text[offset:offset + 3500], parse_mode=None)


def report_text(summary: dict, heading: str = "Сводка журнала") -> str:
    lines = [heading, "Кассовые записи; не бухгалтерская прибыль и не налоговая база."]
    if not summary:
        lines.append("За выбранный период нет действующих записей.")
    for currency, item in sorted(summary.items()):
        lines.extend([f"\n{currency}", f"Поступления: {money_text(item['income'])}",
                      f"Расходы: {money_text(item['expense'])}",
                      f"Разница поступлений и расходов: {money_text(item['balance'])}",
                      f"Записей: {item['count']}"])
    return "\n".join(lines)


def date_range(args: str | None) -> tuple[date | None, date | None]:
    if not args or not args.strip():
        return None, None
    parts = args.split()
    if len(parts) != 2:
        raise ValueError("Нужны две даты: 2026-10-01 2026-10-31")
    try:
        start, end = (date.fromisoformat(x) for x in parts)
    except ValueError:
        raise ValueError("Даты должны быть в формате YYYY-MM-DD") from None
    if start > end:
        raise ValueError("Начало периода позже окончания")
    return start, end


class AccessMiddleware(BaseMiddleware):
    def __init__(self, allowed: frozenset[int]):
        self.allowed = allowed
        self.locks: dict[int, asyncio.Lock] = defaultdict(asyncio.Lock)

    async def __call__(self, handler: Callable[[TelegramObject, dict[str, Any]], Awaitable[Any]],
                       event: TelegramObject, data: dict[str, Any]):
        if not isinstance(event, Message) or event.from_user is None:
            return
        if event.chat.type != "private":
            # Never expose financial information or user IDs in a group.
            return
        user_id = event.from_user.id
        if (event.text or "").split(maxsplit=1)[0:1] == ["/whoami"]:
            await event.answer(f"Ваш Telegram ID: {user_id}", parse_mode=None)
            return
        if user_id not in self.allowed:
            await event.answer("Доступ закрыт. /whoami покажет ID для ALLOWED_USER_IDS.")
            return
        lock = self.locks[user_id]
        if lock.locked():
            await event.answer("Предыдущий запрос ещё обрабатывается. Повторите после ответа.")
            return
        async with lock:
            try:
                return await handler(event, data)
            except ValueError as error:
                await send_text(event, str(error))
            except Exception as error:
                code = secrets.token_hex(4)
                # Exception messages may include keys/URLs or financial content: do not log them.
                LOGGER.error("Request failed code=%s type=%s", code, type(error).__name__)
                await event.answer(
                    f"Не удалось завершить запрос. Код: {code}. "
                    "Проверьте настройки, доступность API/БД и формат файла. "
                    "При сбое API запрос мог быть тарифицирован; автоматического повтора нет.")


class Application:
    def __init__(self, settings: Settings, repo: Repository, ai: AIService):
        self.settings, self.repo, self.ai = settings, repo, ai
        self.router = Router()
        self.selected: dict[int, int | None] = {}
        self.pending_delete: dict[int, tuple[int, str, float]] = {}
        self.bind()

    async def db(self, method: str, *args, **kwargs):
        return await asyncio.to_thread(getattr(self.repo, method), *args, **kwargs)

    async def require_ai(self, owner: int):
        if not self.ai.available:
            raise ValueError("ИИ не настроен. Укажите OPENAI_API_KEY в .env. Локальный учёт доступен.")
        if not await self.db("consent", owner):
            raise ValueError("Данные не отправлены в ИИ. Прочитайте /privacy и введите /consent, "
                             "если разрешаете внешнюю обработку.")

    async def chosen_id(self, owner: int, args: str | None = None) -> int:
        if args and args.strip():
            try:
                ident = int(args.strip())
            except ValueError:
                raise ValueError("Укажите числовой ID документа из /docs") from None
            if ident <= 0:
                raise ValueError("ID должен быть положительным")
            return ident
        if owner in self.selected and self.selected[owner] is not None:
            return self.selected[owner]
        docs = await self.db("list_documents", owner, limit=1)
        if not docs:
            raise ValueError("Нет документов. Пришлите файл или /text данные")
        return docs[0]["id"]

    async def get_doc(self, owner: int, args: str | None = None) -> dict:
        ident = await self.chosen_id(owner, args)
        doc = await self.db("get_document", owner, ident)
        if not doc:
            raise ValueError("Документ не найден среди ваших файлов. Посмотрите /docs")
        return doc

    async def save_source(self, message: Message, source: ParsedSource):
        owner = message.from_user.id
        consent = await self.db("consent", owner)
        if source.images:
            await self.require_ai(owner)
            data = await self.ai.extract(owner, source)
            # Keep native tables, if any, instead of losing deterministic extraction.
            if source.tables:
                data.tables = source.tables + data.tables
                data = DocumentData.model_validate(data.model_dump())
        elif source.tables:
            data = DocumentData(title=source.filename, text=source.text,
                                tables=source.tables, warnings=source.warnings)
        elif consent and self.ai.available and len(source.text) <= 20000:
            data = await self.ai.extract(owner, source)
        else:
            data = DocumentData(title=source.filename, text=source.text, tables=[],
                                warnings=source.warnings + [
                                    "Текст сохранён локально без ИИ. Для таблицы: /table ID после /consent."])
        if source.text and source.text != data.text:
            # Model output must never replace the exact text obtained from native files.
            data.text = source.text + "\n\n--- Дополнение распознавания ИИ ---\n" + data.text
        data = DocumentData.model_validate(data.model_dump())
        ident = await self.db("add_document", owner, source.filename, data)
        self.selected[owner] = ident
        overview = [f"Сохранён документ #{ident}: {data.title[:180]}",
                    "Документ выбран для вопросов.", "В журнал операций ничего не добавлено."]
        for number, table in enumerate(data.tables, 1):
            overview.append(f"Таблица {number}: {table.name[:80]} — {len(table.rows)} строк; столбцы: " +
                            ", ".join(f"{i}. {c[:35]}" for i, c in enumerate(table.columns, 1)))
        if data.warnings:
            overview.append("Предупреждения:\n" + "\n".join(data.warnings))
        overview.append(f"/export {ident} — Excel; /ask ваш вопрос")
        await send_text(message, "\n".join(overview))
        if data.tables:
            payload = await asyncio.to_thread(document_xlsx, data)
            await message.answer_document(BufferedInputFile(payload, filename=f"document_{ident}.xlsx"))

    async def ask(self, message: Message, question: str):
        owner = message.from_user.id
        await self.require_ai(owner)
        if owner in self.selected:
            selected = self.selected[owner]
        else:
            docs = await self.db("list_documents", owner, limit=1)
            selected = docs[0]["id"] if docs else None
        docs = await self.db("documents_for_search", owner, document_id=selected, limit=50)
        if selected is not None and not docs:
            raise ValueError("Выбранный документ удалён. Выберите другой через /use ID")
        summary = await self.db("summary", owner)
        await message.answer("Готовлю ответ. Точные итоги журнала доступны через /report.")
        answer = await self.ai.answer(owner, question, docs, report_text(summary))
        await send_text(message, answer)

    def bind(self):
        r = self.router

        async def table_action(message: Message, ident: str | None = None):
            owner = message.from_user.id
            doc = await self.get_doc(owner, ident)
            data = doc["data"]
            if data.tables:
                payload = await asyncio.to_thread(document_xlsx, data)
                await message.answer_document(
                    BufferedInputFile(payload, filename=f"document_{doc['id']}.xlsx"),
                    caption=f"Таблицы документа #{doc['id']} в Excel."
                )
                return
            await self.require_ai(owner)
            source = ParsedSource(filename=doc["filename"], text=data.text, tables=[],
                                  images=[], warnings=data.warnings)
            await self.save_source(message, source)

        async def chart_action(message: Message, args: str | None):
            try:
                ident, table_number, category_column, value_column = map(
                    int, (args or "").split())
            except ValueError:
                raise ValueError(
                    "Пример: /chart 7 1 1 3 — документ 7, таблица 1, "
                    "столбец категорий 1, числовой столбец 3 (нумерация с 1)."
                ) from None
            doc = await self.get_doc(message.from_user.id, str(ident))
            tables = doc["data"].tables
            if not 1 <= table_number <= len(tables):
                raise ValueError("Неверный номер таблицы. Посмотрите список через /docs и /export ID")
            table = tables[table_number - 1]
            payload = await asyncio.to_thread(
                chart_xlsx, table, category_column - 1, value_column - 1)
            await message.answer_document(
                BufferedInputFile(payload, filename=f"chart_{ident}_{table_number}.xlsx"),
                caption="График и исходные строки в Excel. Проверьте единицы измерения и строки «Итого»."
            )

        @r.message(Command("start", "help", "menu"))
        async def start(message: Message):
            await message.answer(HELP, reply_markup=MAIN_MENU, parse_mode=None)
            await message.answer(
                "ИИ выключен, пока вы не дали согласие. /consent разрешает передавать "
                "внешнему провайдеру вопросы, выбранные фрагменты документов, сводку журнала "
                "и изображения при распознавании. API может быть платным. Подробности: /privacy.",
                reply_markup=MAIN_MENU)

        @r.message(Command("privacy"))
        async def privacy(message: Message):
            await send_text(message,
                "Работа только в личных чатах и только для ID из ALLOWED_USER_IDS. "
                "Каждый пользователь видит свои документы и операции.\n\n"
                "Текст, таблицы, журнал и согласие хранятся в вашей SQLite/PostgreSQL. "
                "Шифрование базы этим кодом не реализовано: защитите диск, резервные копии и доступ к БД. "
                "Оригиналы файлов скачиваются во временную папку и удаляются после обработки; "
                "исходное сообщение/файл остаётся у Telegram. Учтите, что чаты с ботами не являются "
                "сквозным зашифрованным каналом. Не присылайте пароли, карточные данные и секреты.\n\n"
                f"Провайдер ИИ: {self.settings.openai_base_url}; модель: {self.settings.openai_model}. "
                "После /consent ему могут передаваться тексты вопросов, выбранные фрагменты файлов, "
                "сводные суммы журнала и изображения/скан-страницы при загрузке. Ключевые данные "
                "API не отправляются в промпт. store=False не означает отсутствия обработки "
                "или журналов у провайдера; проверьте его условия и договоры.\n\n"
                "/revoke запрещает следующие запросы, но не отзывает ранее переданные данные. "
                "/forget ID удаляет документ из рабочей базы после подтверждения, не из Telegram "
                "и не из ваших резервных копий. /void ID аннулирует запись, сохраняя историю. "
                "Служебные логи не содержат текстов, сумм и сообщений исключений.")

        @r.message(Command("consent"))
        async def consent(message: Message):
            await self.db("set_consent", message.from_user.id, True)
            await message.answer(
                "Согласие сохранено. Следующие операции ИИ могут передавать вопросы, фрагменты "
                "выбранных документов, сводку журнала и изображения внешнему провайдеру из /privacy. "
                "Это может расходовать средства вашего API-аккаунта. Отключение: /revoke.")

        @r.message(Command("revoke"))
        async def revoke(message: Message):
            await self.db("set_consent", message.from_user.id, False)
            await message.answer("ИИ отключён для следующих запросов. Уже отправленные данные не отзываются.")

        @r.message(Command("docs"))
        async def docs(message: Message):
            documents = await self.db("list_documents", message.from_user.id, limit=30)
            await send_text(message, "Последние 30 документов:\n" + (
                "\n".join(f"#{d['id']} — {d['title'][:150]} ({d['filename'][:80]})" for d in documents)
                if documents else "Пока нет документов."))

        @r.message(F.text == MENU_DOCS)
        async def docs_button(message: Message):
            await docs(message)

        @r.message(Command("use"))
        async def use(message: Message, command: CommandObject):
            owner = message.from_user.id
            if (command.args or "").strip().lower() == "all":
                self.selected[owner] = None
                await message.answer("Для вопросов выбраны все ваши документы (до 50). Поиск по фрагментам.")
                return
            if not command.args:
                raise ValueError("Используйте /use ID или /use all")
            doc = await self.get_doc(owner, command.args)
            self.selected[owner] = doc["id"]
            await message.answer(f"Выбран документ #{doc['id']}: {doc['data'].title[:180]}")

        @r.message(Command("text"))
        async def text_source(message: Message, command: CommandObject):
            if not command.args:
                raise ValueError("Напишите /text и исходные данные после пробела")
            source = await asyncio.to_thread(parse_text, command.args, "Текст сообщения")
            await self.save_source(message, source)

        @r.message(F.document | F.photo)
        async def attachment(message: Message, bot: Bot):
            item = message.document if message.document else message.photo[-1]
            filename = message.document.file_name if message.document else "photo.jpg"
            filename = Path((filename or "file").replace("\\", "/")).name[:160]
            metadata = await bot.get_file(item.file_id)
            if (metadata.file_size or 0) > self.settings.max_file_bytes:
                raise ValueError("Максимальный размер файла — 10 МБ. Разделите документ.")
            if message.photo:
                await self.require_ai(message.from_user.id)
            await message.answer("Обрабатываю файл. Для мелкого текста лучше отправлять фото как файл.")
            with tempfile.TemporaryDirectory(prefix="accountant_") as temp:
                path = Path(temp) / ("input" + Path(filename).suffix.lower())
                await bot.download(item.file_id, destination=path, timeout=60)
                if path.stat().st_size > self.settings.max_file_bytes:
                    raise ValueError("Файл больше 10 МБ")
                source = await asyncio.to_thread(parse_file, path, filename)
                # The caption is not parsed into accounting data or treated as instructions.
                await self.save_source(message, source)

        @r.message(Command("ask"))
        async def question(message: Message, command: CommandObject):
            if not command.args:
                raise ValueError("Используйте /ask ваш вопрос; выбор источника: /use ID или /use all")
            await self.ask(message, command.args)

        @r.message(Command("table"))
        async def make_table(message: Message, command: CommandObject):
            await table_action(message, command.args)

        @r.message(Command("export"))
        async def export(message: Message, command: CommandObject):
            doc = await self.get_doc(message.from_user.id, command.args)
            payload = await asyncio.to_thread(document_xlsx, doc["data"])
            await message.answer_document(BufferedInputFile(payload, filename=f"document_{doc['id']}.xlsx"))

        @r.message(Command("chart"))
        async def chart(message: Message, command: CommandObject):
            await chart_action(message, command.args)

        @r.message(Command("sum"))
        async def sum_column(message: Message, command: CommandObject):
            try:
                ident, table_number, col_number = map(int, (command.args or "").split())
            except ValueError:
                raise ValueError("Пример: /sum 7 1 3 — документ 7, таблица 1, столбец 3") from None
            doc = await self.get_doc(message.from_user.id, str(ident))
            tables = doc["data"].tables
            if not 1 <= table_number <= len(tables):
                raise ValueError("Неверный номер таблицы")
            table = tables[table_number - 1]
            if not 1 <= col_number <= len(table.columns):
                raise ValueError("Неверный номер столбца")
            total, count, empty = Decimal("0"), 0, 0
            for row_number, row in enumerate(table.rows, 1):
                cell = row[col_number - 1].strip()
                if not cell:
                    empty += 1
                    continue
                try:
                    value = parse_money(cell)
                except ValueError:
                    raise ValueError(f"Строка {row_number}: значение не является однозначным числом "
                                     "с не более чем 2 знаками после запятой. Сумма не вычислена.") from None
                total += value
                count += 1
            await message.answer(
                f"Документ #{ident}, таблица {table_number}, столбец {col_number}: "
                f"{table.columns[col_number - 1]}\nСумма: {money_text(total)}\n"
                f"Числовых строк: {count}; пустых: {empty}.\n"
                "Это сумма выбранного столбца, включая строки «итого», если они есть. "
                "Валюту/единицу и отсутствие дублей проверьте сами; код их не определяет.")

        @r.message(Command("income", "expense"))
        async def add_entry(message: Message, command: CommandObject):
            parts = (command.args or "").split(maxsplit=3)
            if len(parts) != 4:
                raise ValueError(f"Пример: /{command.command} 1250.50 RUB 2026-10-25 Продажи | Комментарий")
            amount_raw, currency, day_raw, remainder = parts
            try:
                day = date.fromisoformat(day_raw)
            except ValueError:
                raise ValueError("Дата операции: YYYY-MM-DD") from None
            category, separator, note = remainder.partition("|")
            ident, created = await self.db(
                "add_entry", message.from_user.id, command.command.lower(), parse_money(amount_raw),
                currency.upper(), day, category.strip(), note.strip() if separator else "",
                f"telegram:{message.chat.id}:{message.message_id}")
            await message.answer(f"{'Добавлена' if created else 'Уже учтена'} запись #{ident}. "
                                 "Сводка: /report. Аннулировать: /void ID.")

        @r.message(Command("report", "ledger"))
        async def report(message: Message, command: CommandObject):
            start, end = date_range(command.args)
            owner = message.from_user.id
            summary = await self.db("summary", owner, start=start, end=end)
            heading = "Сводка за всё время" if start is None else f"Сводка {start} — {end} включительно"
            await send_text(message, report_text(summary, heading))
            if command.command.lower() == "ledger":
                entries = await self.db("entries", owner, start=start, end=end, limit=10000)
                payload = await asyncio.to_thread(ledger_xlsx, entries, summary)
                await message.answer_document(BufferedInputFile(payload, filename="ledger.xlsx"))

        @r.message(Command("void"))
        async def void(message: Message, command: CommandObject):
            try:
                ident = int(command.args or "")
            except ValueError:
                raise ValueError("Укажите ID записи: /void 12") from None
            changed = await self.db("void_entry", message.from_user.id, ident)
            await message.answer("Запись аннулирована; история сохранена." if changed else
                                 "Запись не найдена или уже аннулирована.")

        @r.message(Command("calc"))
        async def calculate(message: Message, command: CommandObject):
            value = safe_calculate(command.args or "")
            await message.answer(f"Результат: {format(value, 'f')}\nДенежное округление: {money_text(value)}")

        @r.message(Command("vat"))
        async def calculate_vat(message: Message, command: CommandObject):
            parts = (command.args or "").split()
            if len(parts) != 2:
                raise ValueError("Пример: /vat 1200 20; ставка не выбирается автоматически")
            result = vat(parse_money(parts[0]), parse_money(parts[1]))
            await message.answer(
                f"Если сумма включает НДС:\nБаза: {money_text(result['base'])}; "
                f"НДС: {money_text(result['included_tax'])}\n\n"
                f"Если сумма без НДС:\nНДС сверху: {money_text(result['added_tax'])}; "
                f"С НДС: {money_text(result['total_with_tax'])}\nСтавку и применимость определяете вы.")

        @r.message(Command("forget"))
        async def forget(message: Message, command: CommandObject):
            owner = message.from_user.id
            doc = await self.get_doc(owner, command.args)
            token = secrets.token_hex(4)
            self.pending_delete[owner] = (doc["id"], token, time.monotonic() + 120)
            await message.answer(f"Удалить документ #{doc['id']} из рабочей базы? "
                                 "Не удаляет сообщения Telegram, журнал операций и резервные копии.\n"
                                 f"Подтвердите в течение 2 минут: /forget_confirm {token}")

        @r.message(Command("forget_confirm"))
        async def forget_confirm(message: Message, command: CommandObject):
            owner = message.from_user.id
            pending = self.pending_delete.get(owner)
            if not pending or pending[2] < time.monotonic() or not secrets.compare_digest(
                    (command.args or "").strip(), pending[1]):
                raise ValueError("Подтверждение неверно или истекло. Начните заново: /forget ID")
            ident = pending[0]
            deleted = await self.db("delete_document", owner, ident)
            self.pending_delete.pop(owner, None)
            if self.selected.get(owner) == ident:
                self.selected.pop(owner, None)
            await message.answer("Документ удалён." if deleted else "Документ уже отсутствует.")

        @r.message(F.text == MENU_UPLOAD)
        async def upload_button(message: Message):
            await message.answer(
                "Прикрепите сюда фото или файл как документ. Поддерживаются PDF, DOCX, XLSX, "
                "CSV, TXT, JSON, JPG, PNG и WEBP. Фото, сканы и текстовые PDF требуют "
                "настроенного ИИ и согласия /consent."
            )

        @r.message(F.text == MENU_TABLE)
        async def table_button(message: Message):
            await table_action(message)

        @r.message(F.text == MENU_CHART)
        async def chart_button(message: Message):
            doc = await self.get_doc(message.from_user.id)
            if not doc["data"].tables:
                raise ValueError(f"В документе #{doc['id']} таблиц пока нет. Попробуйте /table {doc['id']}.")
            lines = [f"Выбран документ #{doc['id']}. Номера столбцов начинаются с 1."]
            for number, table in enumerate(doc["data"].tables, 1):
                columns = "; ".join(f"{i}. {name}" for i, name in enumerate(table.columns, 1))
                lines.append(f"Таблица {number} «{table.name}»: {columns}")
            lines.append(
                f"Для графика отправьте: /chart {doc['id']} НОМЕР_ТАБЛИЦЫ "
                "СТОЛБЕЦ_КАТЕГОРИИ СТОЛБЕЦ_ЧИСЕЛ\nПример: /chart "
                f"{doc['id']} 1 1 3"
            )
            await message.answer("\n".join(lines))

        @r.message(F.text == MENU_ASK)
        async def ask_button(message: Message):
            await message.answer(
                "Напишите следующим сообщением вопрос по выбранному файлу. "
                "Чтобы выбрать источник, используйте /use ID или /use all."
            )

        @r.message(F.text == MENU_REVIEW)
        async def review_button(message: Message):
            await self.ask(
                message,
                "Проверь выбранный документ или таблицу на возможные пропуски, дубликаты, "
                "повторные заголовки, неоднозначные форматы дат/чисел, несовпадение валют "
                "и единиц. Раздели факты из документа и рекомендации. Укажи источники; "
                "предложи, что можно исправить или уточнить, но ничего не меняй автоматически. "
                "Если контекста недостаточно для проверки всего файла, скажи об этом.",
            )

        @r.message(F.text == MENU_REPORT)
        async def report_button(message: Message):
            summary = await self.db("summary", message.from_user.id)
            await message.answer(report_text(summary))

        @r.message(F.text == MENU_PRIVACY)
        async def privacy_button(message: Message):
            await privacy(message)

        @r.message(F.text == MENU_HELP)
        async def help_button(message: Message):
            await start(message)

        @r.message(F.text & ~F.text.startswith("/"))
        async def normal_question(message: Message):
            await self.ask(message, message.text)

        @r.message()
        async def fallback(message: Message):
            await message.answer("Команда или формат не поддерживается. Список возможностей: /help")


async def main():
    # Applies to files created by this process on Unix. Does not encrypt existing files.
    if os.name == "posix":
        os.umask(0o077)
    logging.basicConfig(level=logging.WARNING, format="%(levelname)s %(name)s %(message)s")
    # Avoid access tokens in debug HTTP request URLs and update logs.
    for name in ("httpx", "httpcore", "openai", "aiogram"):
        logging.getLogger(name).setLevel(logging.CRITICAL)
    settings = Settings.load()
    repo = Repository(settings.database_url)
    await asyncio.to_thread(repo.initialize)
    ai = AIService(settings)
    bot = Bot(settings.bot_token)
    dispatcher = Dispatcher()
    dispatcher.message.outer_middleware(AccessMiddleware(settings.allowed_user_ids))
    application = Application(settings, repo, ai)
    dispatcher.include_router(application.router)
    try:
        await bot.set_my_commands([
            BotCommand(command="menu", description="Открыть кнопочное меню"),
            BotCommand(command="help", description="Команды и примеры"),
            BotCommand(command="docs", description="Мои документы"),
            BotCommand(command="ask", description="Вопрос по выбранным файлам"),
            BotCommand(command="chart", description="Построить график в Excel"),
            BotCommand(command="report", description="Сводка поступлений и расходов"),
            BotCommand(command="ledger", description="Журнал в Excel"),
            BotCommand(command="privacy", description="Хранение и внешняя обработка"),
        ])
        # Does not delete webhooks or discard pending updates automatically.
        print("Бот запущен (polling). Остановить: Ctrl+C.", flush=True)
        await dispatcher.start_polling(bot, allowed_updates=["message"], close_bot_session=False)
    finally:
        await ai.close()
        await bot.session.close()
        await asyncio.to_thread(repo.close)


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except (KeyboardInterrupt, SystemExit):
        pass
    except Exception as error:
        # Do not print possibly sensitive DB URLs / provider errors.
        if isinstance(error, ValueError):
            print(f"Ошибка настроек: {error}")
        else:
            print(f"Не удалось запустить: {type(error).__name__}. Проверьте .env, сеть и БД.")
        raise SystemExit(1) from None
