import os
import sys
import json
from pathlib import Path

APP_NAME = "DocApp"
CONFIG_NAME = "config.json"

# Параметры мероприятия, которые вводятся в окне «Создать шаблоны»
# EVENT_INFO — полное название в именительном падеже,
# EVENT_INFO_PREP / EVENT_INFO_GEN — оно же в предложном и родительном падежах
EVENT_KEYS = ('EVENT_NAME', 'EVENT_INFO', 'EVENT_INFO_PREP', 'EVENT_INFO_GEN',
              'DATE_INFO', 'PLACE_INFO', 'OFERTA_LINK')


def resource_path(rel: str) -> Path:
    """
    Путь к ресурсу, вшитому в exe (read-only), или к файлу в корне проекта
    при запуске из исходников.
    """
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return base / rel

def user_config_dir() -> Path:
    """Где хранить рабочий конфиг (read-write). Без внешних зависимостей."""
    # Windows: %LOCALAPPDATA%\DocApp
    if sys.platform.startswith("win"):
        root = Path(os.getenv("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
    # macOS: ~/Library/Application Support/DocApp
    elif sys.platform == "darwin":
        root = Path.home() / "Library" / "Application Support"
    # Linux: ~/.config/DocApp
    else:
        root = Path(os.getenv("XDG_CONFIG_HOME", Path.home() / ".config"))
    return root / APP_NAME

def load_config() -> dict:
    """Гарантирует наличие рабочего config.json в user_dir и возвращает dict."""
    udir = user_config_dir()
    udir.mkdir(parents=True, exist_ok=True)
    ucfg = udir / CONFIG_NAME

    if not ucfg.exists():
        # 1-й запуск: берём дефолт из ресурсов (если он есть внутри exe)
        default_cfg_path = resource_path(CONFIG_NAME)
        if default_cfg_path.exists():
            ucfg.write_text(default_cfg_path.read_text(encoding="utf-8"), encoding="utf-8")
        else:
            ucfg.write_text(json.dumps({}, ensure_ascii=False, indent=2), encoding="utf-8")

    # читаем рабочий конфиг
    with ucfg.open("r", encoding="utf-8") as f:
        cfg = json.load(f)

    # отсутствующие ключи заполняем пустыми строками, чтобы не падать с KeyError
    for key in EVENT_KEYS:
        cfg.setdefault(key, "")
    return cfg

def require_event_params(params: dict) -> None:
    """Проверяет, что параметры мероприятия заполнены."""
    missing = [k for k in EVENT_KEYS if not str(params.get(k, "")).strip()]
    if missing:
        raise ValueError(
            "Не заполнены параметры мероприятия: " + ", ".join(missing)
            + ". Нажмите «Создать шаблоны» и заполните все поля."
        )

def save_config(cfg: dict) -> None:
    """Сохраняет рабочий конфиг (read-write место)."""
    udir = user_config_dir()
    udir.mkdir(parents=True, exist_ok=True)
    ucfg = udir / CONFIG_NAME
    tmp = ucfg.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(ucfg)
