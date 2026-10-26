from config import Settings


def test_explicit_fadehost_database_url(monkeypatch):
    monkeypatch.setattr("config.load_dotenv", lambda *args, **kwargs: None)
    monkeypatch.setenv("BOT_TOKEN", "123456:TEST_TOKEN_NOT_REAL")
    monkeypatch.setenv("ALLOWED_USER_IDS", "123456789")
    monkeypatch.setenv("DATABASE_URL", "sqlite:////data/accountant.db")
    monkeypatch.setenv("OPENAI_API_KEY", "")
    monkeypatch.setenv("OPENAI_BASE_URL", "https://api.openai.com/v1")
    monkeypatch.setenv("TIMEZONE", "Europe/Moscow")

    settings = Settings.load()

    assert settings.database_url == "sqlite:////data/accountant.db"
    assert settings.allowed_user_ids == frozenset({123456789})
    assert "TEST_TOKEN_NOT_REAL" not in repr(settings)
