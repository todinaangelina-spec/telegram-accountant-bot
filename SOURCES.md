# Документация, проверенная при разработке

- [aiogram: скачивание файлов](https://docs.aiogram.dev/en/latest/api/download_file.html). Используется Bot.download с локальным временным файлом.
- [OpenAI: Structured Outputs](https://developers.openai.com/api/docs/guides/structured-outputs). Извлечение данных в Pydantic-модель через chat.completions.parse; обработка отказов и незавершённых ответов.
- [OpenAI: изображения и vision](https://developers.openai.com/api/docs/guides/images-vision). Изображения передаются в base64 data URL, без публикации файлов. Распознавание может ошибаться.
- [Google Gemini API: OpenAI compatibility](https://ai.google.dev/gemini-api/docs/openai). Описывает совместимые chat completions, structured outputs и обработку изображений; совместимость зависит от модели и поддерживаемых параметров.
- [Google Gemini API: условия использования данных](https://ai.google.dev/gemini-api/terms). Для неоплачиваемого уровня содержимое запросов может использоваться для улучшения сервисов; не отправляйте конфиденциальные документы.
- [SQLAlchemy 2.0: подключения](https://docs.sqlalchemy.org/en/20/core/engines.html). SQLite и PostgreSQL подключаются через DATABASE_URL.

Дата проверки: 25 октября 2026. Доступность конкретной модели OpenAI зависит от вашего API-аккаунта; название задаётся в OPENAI_MODEL.
