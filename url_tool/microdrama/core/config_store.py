from __future__ import annotations

import configparser
from pathlib import Path
from typing import Optional

BASE_DIR = Path(__file__).resolve().parents[2]
CONFIG_FILE = BASE_DIR / "config.ini"
DEFAULT_MODE = 1


def load_mode() -> int:
    """Load the processing mode from the configuration file."""
    config = configparser.ConfigParser()
    if CONFIG_FILE.exists():
        try:
            config.read(CONFIG_FILE, encoding="utf-8")
            mode = config.getint("DEFAULT", "mode", fallback=DEFAULT_MODE)
            if mode in (1, 2):
                return mode
        except Exception:
            pass
    return DEFAULT_MODE


def save_mode(mode: int) -> None:
    """Persist the processing mode to the configuration file."""
    config = configparser.ConfigParser()
    config["DEFAULT"] = {"mode": str(mode)}
    CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
    with CONFIG_FILE.open("w", encoding="utf-8") as config_file:
        config.write(config_file)


def ensure_mode(mode: Optional[int]) -> int:
    """Validate and persist a mode value."""
    if mode not in (1, 2):
        mode = DEFAULT_MODE
    save_mode(mode)
    return mode
