"""Synthetic fixtures for local source ingestion.

The assertions use neutral identifiers and labels rather than financial examples so
that these tests only verify preservation and parser safety behavior.
"""
from __future__ import annotations

from pathlib import Path

import pytest

from ingest import parse_file, parse_text


CSV_CP1251_FIXTURE = "Код;Метка\n0012;Тест\n"
JSON_LIST_FIXTURE = '[{"code":"0012","nested":{"label":"alpha"}},{"code":"0013","nested":["beta",2]}]'


def test_parse_text_is_local_and_requires_no_external_processing() -> None:
    source = parse_text("Простой текст 0012", title="Заметка")

    assert source.filename == "Заметка"
    assert source.text == "Простой текст 0012"
    assert source.tables == []
    assert source.images == []
    assert source.requires_ai is True


def test_csv_sniffs_semicolon_decodes_cp1251_and_preserves_leading_zeroes(tmp_path: Path) -> None:
    path = tmp_path / "fixture.csv"
    path.write_bytes(CSV_CP1251_FIXTURE.encode("cp1251"))

    source = parse_file(path, "display.csv")

    assert source.filename == "display.csv"
    assert source.tables[0].columns == ["Код", "Метка"]
    assert source.tables[0].rows == [["0012", "Тест"]]
    assert source.requires_ai is False


def test_xlsx_uses_first_nonempty_row_and_keeps_formula_as_text(tmp_path: Path) -> None:
    openpyxl = pytest.importorskip("openpyxl")
    path = tmp_path / "fixture.xlsx"
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Синтетика"
    worksheet.append([])
    worksheet.append(["code", "formula"])
    worksheet.append(["0012", "=A3"])
    workbook.save(path)
    workbook.close()

    source = parse_file(path, "fixture.xlsx")

    assert source.tables[0].name == "Синтетика"
    assert source.tables[0].columns == ["code", "formula"]
    assert source.tables[0].rows == [["0012", "=A3"]]
    assert any("первая непустая" in warning for warning in source.warnings)
    assert any("не пересчитывались" in warning for warning in source.warnings)


def test_json_flattens_list_of_dicts_and_serializes_nested_values(tmp_path: Path) -> None:
    path = tmp_path / "fixture.json"
    path.write_text(JSON_LIST_FIXTURE, encoding="utf-8")

    source = parse_file(path, "fixture.json")

    assert source.text == JSON_LIST_FIXTURE
    assert source.tables[0].columns == ["code", "nested"]
    assert source.tables[0].rows == [
        ["0012", '{"label":"alpha"}'],
        ["0013", '["beta",2]'],
    ]


def test_docx_extracts_paragraphs_and_table(tmp_path: Path) -> None:
    docx = pytest.importorskip("docx")
    path = tmp_path / "fixture.docx"
    document = docx.Document()
    document.add_paragraph("Перед таблицей")
    table = document.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "code"
    table.cell(0, 1).text = "label"
    table.cell(1, 0).text = "0012"
    table.cell(1, 1).text = "gamma"
    document.add_paragraph("После таблицы")
    document.save(path)

    source = parse_file(path, "fixture.docx")

    assert "Перед таблицей" in source.text
    assert "[Таблица 1]" in source.text
    assert "После таблицы" in source.text
    assert source.tables[0].columns == ["code", "label"]
    assert source.tables[0].rows == [["0012", "gamma"]]
    assert any("Колонтитулы" in warning for warning in source.warnings)


def test_blank_pdf_has_page_marker_and_raster_image(tmp_path: Path) -> None:
    pypdf = pytest.importorskip("pypdf")
    path = tmp_path / "fixture.pdf"
    writer = pypdf.PdfWriter()
    writer.add_blank_page(width=144, height=144)
    with path.open("wb") as output:
        writer.write(output)

    source = parse_file(path, "fixture.pdf")

    assert "[Страница 1]" in source.text
    assert len(source.images) == 1
    assert source.images[0][0] == "image/png"
    assert source.requires_ai is True
    assert any("OCR или модель" in warning for warning in source.warnings)


def test_legacy_and_unknown_extensions_explain_supported_conversion(tmp_path: Path) -> None:
    legacy = tmp_path / "fixture.xls"
    legacy.write_bytes(b"synthetic")
    unknown = tmp_path / "fixture.bin"
    unknown.write_bytes(b"synthetic")

    with pytest.raises(ValueError, match=r"конвертируйте.*xlsx"):
        parse_file(legacy, "fixture.xls")
    with pytest.raises(ValueError, match="Неподдерживаемый формат"):
        parse_file(unknown, "fixture.bin")


def test_csv_row_cap_rejects_instead_of_silently_truncating(tmp_path: Path) -> None:
    path = tmp_path / "over-limit.csv"
    path.write_text("code\n" + "\n".join(f"row-{index}" for index in range(2001)), encoding="utf-8")

    with pytest.raises(ValueError, match="более 2000 строк"):
        parse_file(path, "over-limit.csv")


def test_text_cap_rejects_instead_of_silently_truncating() -> None:
    with pytest.raises(ValueError, match="более 200000 символов"):
        parse_text("x" * 200_001)


def test_json_decimal_preserved_without_float_rounding(tmp_path):
    path = tmp_path / "precision.json"
    original = '[{"amount":0.10000000000000000001,"nested":{"value":0.10000000000000000001}}]'
    path.write_text(original, encoding="utf-8")
    source = parse_file(path, path.name)
    assert source.text == original
    assert source.tables[0].rows[0][0] == "0.10000000000000000001"
    assert "0.10000000000000000001" in source.tables[0].rows[0][1]


def test_xlsx_huge_leading_blank_range_rejected(tmp_path):
    from openpyxl import Workbook
    path = tmp_path / "range.xlsx"
    workbook = Workbook()
    workbook.active["A50000"] = "code"
    workbook.save(path)
    workbook.close()
    with pytest.raises(ValueError, match="2001"):
        parse_file(path, path.name)


def test_pdf_images_have_original_page_mapping(tmp_path):
    from pypdf import PdfWriter
    path = tmp_path / "pages.pdf"
    writer = PdfWriter()
    writer.add_blank_page(width=72, height=72)
    writer.add_blank_page(width=72, height=72)
    with path.open("wb") as stream:
        writer.write(stream)
    source = parse_file(path, path.name)
    assert any("Изображение 2 соответствует странице 2" in x for x in source.warnings)
