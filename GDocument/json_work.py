import json
from pathlib import Path

CONFIG_PATH = Path("config.json")

def save_config_json(params: dict, filename: str = None):
    """Сохраняет словарь параметров в JSON."""
    path = Path(filename) if filename else CONFIG_PATH
    with open(path, "w", encoding="utf-8") as f:
        json.dump(params, f, ensure_ascii=False, indent=2)

def load_config_json(filename: str = None, defaults: dict = None) -> dict:
    """Загружает параметры из JSON. Если файла нет — возвращает defaults или {}."""
    path = Path(filename) if filename else CONFIG_PATH
    if not path.exists():
        return dict(defaults) if defaults else {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)
