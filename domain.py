"""Shared, strictly validated document models."""
from pydantic import BaseModel, ConfigDict, Field, model_validator


class TableData(BaseModel):
    model_config = ConfigDict(extra="forbid")
    name: str
    columns: list[str]
    rows: list[list[str]]

    @model_validator(mode="after")
    def validate_table(self):
        if not self.columns or len(self.columns) > 50:
            raise ValueError("Таблица должна содержать от 1 до 50 столбцов")
        if len(self.rows) > 2000:
            raise ValueError("В одной таблице допускается не более 2000 строк")
        if any(len(row) != len(self.columns) for row in self.rows):
            raise ValueError("Количество ячеек не совпадает с количеством столбцов")
        return self


class DocumentData(BaseModel):
    model_config = ConfigDict(extra="forbid")
    title: str
    text: str
    tables: list[TableData]
    warnings: list[str]

    @model_validator(mode="after")
    def validate_document(self):
        if len(self.tables) > 10:
            raise ValueError("Не более 10 таблиц в документе")
        if sum(len(t.rows) * len(t.columns) for t in self.tables) > 30000:
            raise ValueError("Документ содержит более 30000 ячеек")
        if len(self.text) > 200000:
            raise ValueError("Текст документа превышает 200000 символов")
        return self
