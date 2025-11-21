"""Configuration utilities and constants for BarTender GUI app."""
from __future__ import annotations

import json
import os
import sys
from typing import Any, Dict

APP_TITLE = "BarTender GUI V2.0 batch"
APP_VERSION = "2.0"
PREVIEW_NAME = "preview.png"


def _app_base_dir() -> str:
    """Return the directory where the app should look for resources."""
    try:
        if getattr(sys, "frozen", False):
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        return os.getcwd()


BASE_DIR = _app_base_dir()
PRODUCT_MAP_DEFAULT = os.path.join(BASE_DIR, "Список товаров.xlsx")


def _cfg_dir() -> str:
    base = os.path.join(os.environ.get("APPDATA", os.getcwd()), "BarTenderGUI")
    os.makedirs(base, exist_ok=True)
    return base


def _cfg_path() -> str:
    return os.path.join(_cfg_dir(), "config.json")


def load_config() -> Dict[str, Any]:
    p = _cfg_path()
    if os.path.isfile(p):
        try:
            with open(p, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_config(cfg: Dict[str, Any]) -> None:
    try:
        with open(_cfg_path(), "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
