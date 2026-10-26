"""Application settings; never ship actual credentials."""
import os
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import urlparse
from zoneinfo import ZoneInfo

from dotenv import load_dotenv


@dataclass(frozen=True)
class Settings:
    bot_token: str = field(repr=False)
    allowed_user_ids: frozenset[int]
    database_url: str = field(repr=False)
    openai_api_key: str = field(repr=False, default="")
    openai_base_url: str = "https://api.openai.com/v1"
    openai_model: str = "gpt-4.1-mini"
    timezone: str = "Europe/Moscow"
    max_file_bytes: int = 10 * 1024 * 1024
    ai_concurrency: int = 2

    @classmethod
    def load(cls):
        root = Path(__file__).resolve().parent
        # This project reads only its own .env. Existing OS variables take precedence.
        load_dotenv(root / ".env", override=False)
        token = os.getenv("BOT_TOKEN", "").strip()
        if not token or token == "PASTE_TELEGRAM_BOT_TOKEN":
            raise ValueError("Укажите BOT_TOKEN в .env (токен от @BotFather)")
        raw = os.getenv("ALLOWED_USER_IDS", "")
        try:
            allowed = frozenset(int(x.strip()) for x in raw.split(",") if x.strip())
        except ValueError:
            raise ValueError("ALLOWED_USER_IDS: числовые Telegram ID через запятую") from None
        if not allowed or any(x <= 0 for x in allowed):
            raise ValueError("Заполните ALLOWED_USER_IDS. Доступ всем запрещён по умолчанию.")
        database = os.getenv("DATABASE_URL", "").strip()
        if not database:
            # FadeHost replaces the code directory on each Git deploy. It exposes
            # /data as persistent writable storage, so prefer that location when
            # present; on a local computer use ./data beside this file instead.
            persistent_dir = Path("/data")
            if persistent_dir.is_dir() and os.access(persistent_dir, os.W_OK):
                data_dir = persistent_dir
            else:
                data_dir = root / "data"
            data_dir.mkdir(mode=0o700, exist_ok=True)
            database = f"sqlite:///{data_dir / 'accountant.db'}"
        if not database.startswith(("sqlite:", "sqlite+pysqlite:", "postgresql+psycopg:")):
            raise ValueError("Поддерживаются sqlite:///... или postgresql+psycopg://...")
        base = os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1").strip().rstrip("/")
        parsed = urlparse(base)
        if parsed.scheme != "https" or not parsed.netloc or parsed.username or parsed.password:
            raise ValueError("OPENAI_BASE_URL должен быть HTTPS URL без логина/пароля")
        zone = os.getenv("TIMEZONE", "Europe/Moscow").strip()
        ZoneInfo(zone)
        key = os.getenv("OPENAI_API_KEY", "").strip()
        if key == "PASTE_OPENAI_API_KEY":
            key = ""
        return cls(token, allowed, database, key, base,
                   os.getenv("OPENAI_MODEL", "gpt-4.1-mini").strip(), zone)
