import json
from pathlib import Path

# CONFIG_PATH = Path("config.json")

# def save_config_json(params: dict, filename: str = None):
#     """Сохраняет словарь параметров в JSON."""
#     path = Path(filename) if filename else CONFIG_PATH
#     with open(path, "w", encoding="utf-8") as f:
#         json.dump(params, f, ensure_ascii=False, indent=2)

# def load_config_json(filename: str = None, defaults: dict = None) -> dict:
#     """Загружает параметры из JSON. Если файла нет — возвращает defaults или {}."""
#     path = Path(filename) if filename else CONFIG_PATH
#     if not path.exists():
#         return dict(defaults) if defaults else {}
#     with open(path, "r", encoding="utf-8") as f:
#         return json.load(f)
    
import os

import sys, json
from pathlib import Path

APP_NAME = "DocApp"
CONFIG_NAME = "config.json"

def resource_path(rel: str) -> Path:
    """Файлы, вшитые в exe (read-only): дефолтный config.json, картинки и т.п."""
    base = Path(getattr(sys, "_MEIPASS", Path.cwd()))
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
            # или создаём пустой/минимальный конфиг, если дефолта нет
            ucfg.write_text(json.dumps({"version": 1}, ensure_ascii=False, indent=2), encoding="utf-8")

    # читаем рабочий конфиг
    with ucfg.open("r", encoding="utf-8") as f:
        return json.load(f)

def save_config(cfg: dict) -> None:
    """Сохраняет рабочий конфиг (read-write место)."""
    udir = user_config_dir()
    udir.mkdir(parents=True, exist_ok=True)
    ucfg = udir / CONFIG_NAME
    tmp = ucfg.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(ucfg)

