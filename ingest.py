"""Safe, loss-conscious ingestion for user supplied document files.

This module only extracts local file content.  It intentionally never invokes OCR,
models, or financial calculations; callers may use ``ParsedSource.requires_ai`` to
decide whether a separately-authorized extraction step is appropriate.
"""
from __future__ import annotations

import csv
import io
import json
import re
import zipfile
from dataclasses import dataclass
from datetime import date, datetime, time
from decimal import Decimal
from pathlib import Path
from typing import Any, Iterable

from domain import TableData


MAX_INPUT_BYTES = 10 * 1024 * 1024
MAX_TEXT_CHARS = 200_000
MAX_TABLES = 10
MAX_ROWS_PER_TABLE = 2_000
MAX_COLUMNS = 50
MAX_TOTAL_CELLS = 30_000
MAX_ARCHIVE_DECOMPRESSED_BYTES = 50 * 1024 * 1024
MAX_ARCHIVE_MEMBER_BYTES = 20 * 1024 * 1024
MAX_ARCHIVE_MEMBERS = 2_000
MAX_ARCHIVE_COMPRESSION_RATIO = 100
MAX_PDF_PAGES = 20
MAX_PDF_IMAGE_PAGES = 10
MAX_IMAGE_PIXELS = 25_000_000


@dataclass
class ParsedSource:
    """Normalized local source content, ready for conversion to ``DocumentData``."""

    filename: str
    text: str
    tables: list[TableData]
    images: list[tuple[str, bytes]]
    warnings: list[str]

    @property
    def requires_ai(self) -> bool:
        """Whether the source has images or no structured table content."""
        return bool(self.images) or not self.tables


def _limit_error(limit: str) -> ValueError:
    return ValueError(f"{limit}. Разделите файл на части и повторите загрузку.")


def _check_input_size(path: Path) -> None:
    try:
        size = path.stat().st_size
    except OSError as exc:
        raise ValueError(f"Не удалось прочитать файл: {path.name}") from exc
    if size > MAX_INPUT_BYTES:
        raise _limit_error("Размер файла превышает 10 МБ")


def _check_text(text: str) -> str:
    if len(text) > MAX_TEXT_CHARS:
        raise _limit_error("Текст содержит более 200000 символов")
    return text


def _check_archive_safety(path: Path, label: str) -> None:
    """Inspect ZIP metadata before a DOCX/XLSX library parses the archive."""
    if not zipfile.is_zipfile(path):
        raise ValueError(f"{label} должен быть корректным ZIP-архивом")
    try:
        with zipfile.ZipFile(path) as archive:
            infos = archive.infolist()
    except (OSError, zipfile.BadZipFile) as exc:
        raise ValueError(f"Не удалось безопасно открыть {label}") from exc

    if len(infos) > MAX_ARCHIVE_MEMBERS:
        raise _limit_error("Архив содержит более 2000 элементов")

    total_size = 0
    for info in infos:
        if info.flag_bits & 0x1:
            raise ValueError(f"{label} содержит зашифрованный ZIP-элемент")
        if info.file_size > MAX_ARCHIVE_MEMBER_BYTES:
            raise _limit_error("Один элемент архива превышает 20 МБ после распаковки")
        total_size += info.file_size
        if total_size > MAX_ARCHIVE_DECOMPRESSED_BYTES:
            raise _limit_error("Размер архива после распаковки превышает 50 МБ")
        if info.file_size:
            if not info.compress_size:
                raise ValueError("Архив имеет небезопасную степень сжатия")
            if info.file_size / info.compress_size > MAX_ARCHIVE_COMPRESSION_RATIO:
                raise ValueError("Архив имеет небезопасную степень сжатия")


def _text_cell(value: Any) -> str:
    """Represent a cell without applying accounting/number formatting changes."""
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (datetime, date, time)):
        return value.isoformat(sep=" ") if isinstance(value, datetime) else value.isoformat()
    return str(value)


def _has_value(value: Any) -> bool:
    return value is not None and _text_cell(value) != ""


def _trim_trailing_empty(row: Iterable[Any]) -> list[Any]:
    values = list(row)
    while values and not _has_value(values[-1]):
        values.pop()
    return values


def _make_columns(header: list[Any], width: int, warnings: list[str], table_name: str) -> list[str]:
    columns: list[str] = []
    seen: set[str] = set()
    changed = False
    for index in range(width):
        raw = _text_cell(header[index]) if index < len(header) else ""
        candidate = raw if raw.strip() else f"Столбец {index + 1}"
        if candidate != raw:
            changed = True
        base = candidate
        suffix = 2
        while candidate in seen:
            candidate = f"{base} ({suffix})"
            suffix += 1
            changed = True
        seen.add(candidate)
        columns.append(candidate)
    if changed:
        warnings.append(
            f"В таблице «{table_name}» пустые или повторяющиеся заголовки были "
            "заменены уникальными именами столбцов."
        )
    return columns


def _table_from_rows(
    name: str,
    source_rows: Iterable[Iterable[Any]],
    warnings: list[str],
    *,
    header_first_nonempty: bool = False,
    trim_trailing_cells: bool = True,
) -> TableData | None:
    """Build one rectangular ``TableData`` and reject, rather than truncate, excess."""
    rows = [
        _trim_trailing_empty(row) if trim_trailing_cells else list(row)
        for row in source_rows
    ]
    while rows and not any(_has_value(value) for value in rows[-1]):
        rows.pop()
    if not rows:
        return None

    if header_first_nonempty:
        header_index = next(
            (idx for idx, row in enumerate(rows) if any(_has_value(value) for value in row)),
            None,
        )
        if header_index is None:
            return None
        warnings.append(
            f"В таблице «{name}» первая непустая строка выбрана как заголовок."
        )
        header = rows[header_index]
        data_source = rows[header_index + 1 :]
    else:
        header = rows[0]
        data_source = rows[1:]

    if len(data_source) > MAX_ROWS_PER_TABLE:
        raise _limit_error(f"Таблица «{name}» содержит более 2000 строк")

    width = max([len(header), *(len(row) for row in data_source)], default=0)
    if width == 0:
        return None
    if width > MAX_COLUMNS:
        raise _limit_error(f"Таблица «{name}» содержит более 50 столбцов")

    columns = _make_columns(header, width, warnings, name)
    padded = False
    data: list[list[str]] = []
    for source_row in data_source:
        values = [_text_cell(value) for value in source_row]
        if len(values) < width:
            values.extend([""] * (width - len(values)))
            padded = True
        data.append(values)
    if padded:
        warnings.append(f"Строки таблицы «{name}» дополнены пустыми ячейками до ширины заголовка.")

    return TableData(name=name, columns=columns, rows=data)


def _validate_tables(tables: list[TableData]) -> None:
    if len(tables) > MAX_TABLES:
        raise _limit_error("Документ содержит более 10 таблиц")
    cells = sum(len(table.rows) * len(table.columns) for table in tables)
    if cells > MAX_TOTAL_CELLS:
        raise _limit_error("Документ содержит более 30000 ячеек")


def _decode_text_bytes(data: bytes, kind: str) -> str:
    for encoding in ("utf-8-sig", "cp1251"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise ValueError(f"{kind} должен быть в кодировке UTF-8 или Windows-1251")


def parse_text(text: str, title: str = "Текст") -> ParsedSource:
    """Create a plain-text source without any AI or external operation."""
    if not isinstance(text, str):
        raise TypeError("text должен быть строкой")
    return ParsedSource(filename=title, text=_check_text(text), tables=[], images=[], warnings=[])


def _parse_csv(path: Path, filename: str) -> ParsedSource:
    text = _decode_text_bytes(path.read_bytes(), "CSV")
    try:
        dialect = csv.Sniffer().sniff(text[:8192], delimiters=",;\t")
        delimiter = dialect.delimiter
    except csv.Error:
        delimiter = next((mark for mark in (",", ";", "\t") if mark in text), ",")

    try:
        rows = list(csv.reader(io.StringIO(text, newline=""), delimiter=delimiter))
    except csv.Error as exc:
        raise ValueError("CSV содержит некорректные кавычки или строки") from exc
    warnings: list[str] = []
    table = _table_from_rows(
        Path(filename).stem or "CSV",
        rows,
        warnings,
        trim_trailing_cells=False,
    )
    tables = [table] if table is not None else []
    _validate_tables(tables)
    return ParsedSource(filename=filename, text="", tables=tables, images=[], warnings=warnings)


def _json_load_strict(text: str) -> Any:
    def reject_constant(value: str) -> Any:
        raise ValueError(f"Недопустимая JSON-константа: {value}")

    def object_without_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                raise ValueError(f"Повторяющийся ключ JSON: {key}")
            result[key] = value
        return result

    try:
        return json.loads(
            text,
            parse_float=Decimal,
            parse_constant=reject_constant,
            object_pairs_hook=object_without_duplicate_keys,
        )
    except (json.JSONDecodeError, ValueError) as exc:
        raise ValueError(f"JSON должен иметь строгий корректный синтаксис: {exc}") from exc


def _json_value_text(value: Any) -> str:
    if isinstance(value, str):
        return value
    if value is None:
        return "null"
    if value is True:
        return "true"
    if value is False:
        return "false"
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False, separators=(",", ":"), allow_nan=False, default=str)
    return str(value)


def _parse_json(path: Path, filename: str) -> ParsedSource:
    original = _check_text(_decode_text_bytes(path.read_bytes(), "JSON"))
    parsed = _json_load_strict(original)
    warnings: list[str] = ["Исходный JSON сохранён. Вложенные десятичные числа в ячейках представлены строками без потери точности."]
    if isinstance(parsed, list) and parsed and all(isinstance(item, dict) for item in parsed):
        headers: list[str] = []
        for item in parsed:
            for key in item:
                if key not in headers:
                    headers.append(key)
        source_rows: list[list[Any]] = [headers]
        source_rows.extend([[item.get(header, "") for header in headers] for item in parsed])
        # JSON values need compact serialization for nested values, booleans, and null.
        converted_rows = [[_json_value_text(value) for value in row] for row in source_rows]
        table = _table_from_rows(
            Path(filename).stem or "JSON",
            converted_rows,
            warnings,
            trim_trailing_cells=False,
        )
        tables = [table] if table is not None else []
        _validate_tables(tables)
        return ParsedSource(
            filename=filename,
            text=original,
            tables=tables,
            images=[],
            warnings=warnings,
        )
    return ParsedSource(
        filename=filename,
        text=_check_text(original),
        tables=[],
        images=[],
        warnings=warnings,
    )


def _parse_xlsx(path: Path, filename: str) -> ParsedSource:
    _check_archive_safety(path, "XLSX")
    try:
        from openpyxl import load_workbook
    except ImportError as exc:  # pragma: no cover - dependency declared by the project
        raise RuntimeError("Для XLSX нужен пакет openpyxl") from exc

    try:
        workbook = load_workbook(path, read_only=True, data_only=False)
    except Exception as exc:
        raise ValueError("Не удалось прочитать XLSX") from exc

    warnings: list[str] = ["Числовые форматы XLSX не воспроизводятся: ведущие нули сохраняются в текстовых ячейках, но не в форматах вроде 00000."]
    tables: list[TableData] = []
    formulas_found = False
    try:
        if len(workbook.worksheets) > MAX_TABLES:
            raise _limit_error("XLSX содержит более 10 листов")
        for worksheet in workbook.worksheets:
            if (worksheet.max_column or 0) > MAX_COLUMNS or (worksheet.max_row or 0) > MAX_ROWS_PER_TABLE + 1:
                raise _limit_error(f"Лист «{worksheet.title}» имеет размер более 2001 строк или 50 столбцов (включая пустые/форматированные)")
            # Ignore underspecified dimensions so actual cells are not silently skipped.
            worksheet.reset_dimensions()
            raw_rows: list[list[Any]] = []
            header_seen = False
            data_rows_seen = 0
            for row in worksheet.iter_rows():
                if len(raw_rows) >= MAX_ROWS_PER_TABLE + 1 or len(row) > MAX_COLUMNS:
                    raise _limit_error(f"Лист «{worksheet.title}» превышает 2001 строку или 50 столбцов, включая пустые")
                raw: list[Any] = []
                for cell in row:
                    if cell.data_type == "f":
                        formulas_found = True
                    raw.append(cell.value)
                raw_rows.append(raw)
                if not header_seen and any(_has_value(value) for value in raw):
                    header_seen = True
                elif header_seen:
                    data_rows_seen += 1
                # Count rows only after the first nonempty header. This prevents a huge
                # worksheet from consuming memory while not penalizing leading blank rows.
                if data_rows_seen > MAX_ROWS_PER_TABLE:
                    raise _limit_error(f"Лист «{worksheet.title}» содержит более 2000 строк")
            table = _table_from_rows(
                worksheet.title,
                raw_rows,
                warnings,
                header_first_nonempty=True,
            )
            if table is not None:
                tables.append(table)
        _validate_tables(tables)
    finally:
        workbook.close()

    if formulas_found:
        warnings.append("Формулы XLSX сохранены как текст и не пересчитывались.")
    return ParsedSource(filename=filename, text="", tables=tables, images=[], warnings=warnings)


def _parse_docx(path: Path, filename: str) -> ParsedSource:
    _check_archive_safety(path, "DOCX")
    try:
        from docx import Document
        from docx.table import Table
        from docx.text.paragraph import Paragraph
    except ImportError as exc:  # pragma: no cover - dependency declared by the project
        raise RuntimeError("Для DOCX нужен пакет python-docx") from exc

    try:
        document = Document(path)
    except Exception as exc:
        raise ValueError("Не удалось прочитать DOCX") from exc

    if len(document.tables) > MAX_TABLES:
        raise _limit_error("DOCX содержит более 10 таблиц")

    warnings: list[str] = [
        "DOCX: читается основной текст и обычные таблицы. Колонтитулы, сноски, текстовые поля и вложенные изображения не извлекаются."
    ]
    text_parts: list[str] = []
    tables: list[TableData] = []
    table_number = 0

    # The document body order covers ordinary paragraphs and tables.  Headers, footers,
    # drawing text, and other complex elements are deliberately not treated as paragraphs.
    body = document.element.body
    for child in body.iterchildren():
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            paragraph = Paragraph(child, document)
            if paragraph.text:
                text_parts.append(paragraph.text)
        elif tag == "tbl":
            table_number += 1
            docx_table = Table(child, document)
            raw_rows = [[cell.text for cell in row.cells] for row in docx_table.rows]
            table_name = f"Таблица {table_number}"
            table = _table_from_rows(
                table_name,
                raw_rows,
                warnings,
                trim_trailing_cells=False,
            )
            if table is not None:
                tables.append(table)
            text_parts.append(f"[Таблица {table_number}]")

    _validate_tables(tables)
    text = _check_text("\n".join(text_parts))
    return ParsedSource(filename=filename, text=text, tables=tables, images=[], warnings=warnings)


def _page_has_embedded_image(page: Any) -> bool:
    """Check image XObjects without decoding their data."""
    try:
        resources = page.get("/Resources")
        if resources is None:
            return False
        resources = resources.get_object()
        xobjects = resources.get("/XObject")
        if xobjects is None:
            return False
        for xobject in xobjects.get_object().values():
            obj = xobject.get_object()
            if obj.get("/Subtype") == "/Image":
                return True
    except Exception:
        # Failing closed would raster every exotic PDF page and could reject a useful file.
        # pypdf text detection remains the safe fallback.
        return False
    return False


def _meaningful_text(text: str) -> str:
    # Character count excludes whitespace so blank/near-blank OCR layers rasterize.
    return re.sub(r"\s+", "", text)


def _render_pdf_page(pdf_document: Any, index: int, warnings: list[str]) -> bytes:
    """Raster one page at roughly 160 DPI without exceeding a 2500px long edge."""
    page = pdf_document[index]
    try:
        width, height = page.get_size()
        base_scale = 160 / 72
        edge_scale = 2500 / max(width, height) if max(width, height) else base_scale
        scale = min(base_scale, edge_scale)
        bitmap = page.render(scale=scale)
        try:
            image = bitmap.to_pil().convert("RGB")
        finally:
            bitmap.close()
        if max(image.size) > 2500:
            # PDFium rounding can put the rendered edge one pixel over the calculated cap.
            # This is a downscale of a generated raster, never a crop or enlargement.
            image.thumbnail((2500, 2500))
        output = io.BytesIO()
        image.save(output, format="PNG", optimize=False)
        if scale < base_scale:
            warnings.append(
                f"Растеризованная страница {index + 1} уменьшена, чтобы не превышать 2500 пикселей по длинной стороне."
            )
        return output.getvalue()
    finally:
        page.close()


def _parse_pdf(path: Path, filename: str) -> ParsedSource:
    try:
        from pypdf import PdfReader
        import pypdfium2 as pdfium
    except ImportError as exc:  # pragma: no cover - dependencies declared by the project
        raise RuntimeError("Для PDF нужны пакеты pypdf и pypdfium2") from exc

    try:
        reader = PdfReader(path)
    except Exception as exc:
        raise ValueError("Не удалось прочитать PDF") from exc
    if reader.is_encrypted:
        raise ValueError("PDF защищён паролем; загрузите незашифрованную копию.")
    if len(reader.pages) > MAX_PDF_PAGES:
        raise _limit_error("PDF содержит более 20 страниц")

    page_texts: list[str] = []
    pages_to_render: list[int] = []
    for index, page in enumerate(reader.pages):
        try:
            extracted = page.extract_text() or ""
        except Exception:
            extracted = ""
        page_texts.append(f"[Страница {index + 1}]\n{extracted}")
        if len(_meaningful_text(extracted)) < 80 or _page_has_embedded_image(page):
            pages_to_render.append(index)

    text = _check_text("\n\n".join(page_texts))
    if len(pages_to_render) > MAX_PDF_IMAGE_PAGES:
        raise _limit_error("PDF требует растеризации более 10 страниц")

    warnings: list[str] = []
    images: list[tuple[str, bytes]] = []
    if pages_to_render:
        try:
            pdf_document = pdfium.PdfDocument(str(path))
        except Exception as exc:
            raise ValueError("Не удалось растеризовать PDF") from exc
        try:
            for index in pages_to_render:
                image_data = _render_pdf_page(pdf_document, index, warnings)
                images.append(("image/png", image_data))
                warnings.append(f"Изображение {len(images)} соответствует странице {index + 1} PDF.")
                warnings.append(
                    f"Страница {index + 1} PDF преобразована в изображение; OCR или модель могут допустить ошибки."
                )
        finally:
            pdf_document.close()

    return ParsedSource(filename=filename, text=text, tables=[], images=images, warnings=warnings)


def _parse_image(path: Path, filename: str) -> ParsedSource:
    try:
        from PIL import Image, ImageOps, UnidentifiedImageError
    except ImportError as exc:  # pragma: no cover - dependency declared by the project
        raise RuntimeError("Для изображений нужен пакет Pillow") from exc

    raw = path.read_bytes()
    try:
        with Image.open(io.BytesIO(raw)) as probe:
            image_format = probe.format
            width, height = probe.size
            if width * height > MAX_IMAGE_PIXELS:
                raise _limit_error("Изображение содержит более 25 миллионов пикселей")
            probe.verify()
    except UnidentifiedImageError as exc:
        raise ValueError("Файл не является корректным изображением PNG, JPEG или WebP") from exc
    except OSError as exc:
        raise ValueError("Не удалось проверить изображение") from exc

    if image_format not in {"PNG", "JPEG", "WEBP"}:
        raise ValueError("Поддерживаются только изображения PNG, JPEG и WebP")

    try:
        with Image.open(io.BytesIO(raw)) as original:
            orientation = original.getexif().get(274, 1)
            normalized = ImageOps.exif_transpose(original)
            if normalized.width * normalized.height > MAX_IMAGE_PIXELS:
                raise _limit_error("Изображение содержит более 25 миллионов пикселей")
            can_keep_original = (
                image_format in {"PNG", "JPEG"}
                and original.mode == "RGB"
                and orientation in (None, 1)
            )
            if can_keep_original:
                mime = "image/png" if image_format == "PNG" else "image/jpeg"
                data = raw
                warnings: list[str] = []
            else:
                # PNG retains pixel content after orientation/color normalization; no resize,
                # crop, or lossy recompression is performed here.
                rgb = normalized.convert("RGB")
                output = io.BytesIO()
                rgb.save(output, format="PNG", optimize=False)
                mime = "image/png"
                data = output.getvalue()
                warnings = ["Изображение нормализовано в RGB PNG без обрезки или увеличения."]
    except OSError as exc:
        raise ValueError("Не удалось обработать изображение") from exc

    return ParsedSource(filename=filename, text="", tables=[], images=[(mime, data)], warnings=warnings)


def parse_file(path: Path, filename: str) -> ParsedSource:
    """Parse one permitted local file into text, tables, and/or image bytes.

    ``filename`` is retained for user-facing metadata; the real path extension chooses
    the parser so a display name cannot cause a binary file to be misclassified.
    """
    path = Path(path)
    if not path.is_file():
        raise ValueError(f"Файл не найден: {path}")
    _check_input_size(path)
    suffix = path.suffix.lower()
    if suffix in {".xls", ".doc"}:
        raise ValueError(f"Формат {suffix} не поддерживается; сначала конвертируйте файл в .xlsx или .docx.")
    if suffix == ".txt":
        return parse_text(_decode_text_bytes(path.read_bytes(), "Текстовый файл"), title=filename)
    if suffix == ".csv":
        return _parse_csv(path, filename)
    if suffix == ".json":
        return _parse_json(path, filename)
    if suffix == ".xlsx":
        return _parse_xlsx(path, filename)
    if suffix == ".docx":
        return _parse_docx(path, filename)
    if suffix == ".pdf":
        return _parse_pdf(path, filename)
    if suffix in {".png", ".jpg", ".jpeg", ".webp"}:
        return _parse_image(path, filename)
    raise ValueError(
        "Неподдерживаемый формат. Допустимы .pdf, .docx, .xlsx, .csv, .txt, .json, "
        ".png, .jpg, .jpeg, .webp; .xls и .doc сначала конвертируйте."
    )
