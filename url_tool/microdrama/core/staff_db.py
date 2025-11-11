from __future__ import annotations

import json
from pathlib import Path
from typing import Dict, Optional

from ..utils.logger import get_logger

LOGGER = get_logger(__name__)

BASE_DIR = Path(__file__).resolve().parents[2]
DATA_DIR = BASE_DIR / "data"
STAFF_DB_FILE = DATA_DIR / "staff_database.json"

DEFAULT_STAFF: Dict[str, str] = {
    "梁应伟": "001X",
    "邹林伶": "5829",
    "赵志强": "7299",
    "杨华": "4241",
    "廖政": "1610",
    "万亭": "174X",
    "任雪梅": "5802",
    "冉小娟": "1363",
    "张静": "8525",
}


def ensure_data_dir() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)


def load_staff_database() -> Dict[str, str]:
    ensure_data_dir()
    if STAFF_DB_FILE.exists():
        try:
            with STAFF_DB_FILE.open("r", encoding="utf-8") as handle:
                data = json.load(handle)
            if isinstance(data, dict):
                return {str(k): str(v) for k, v in data.items()}
        except Exception:  # pragma: no cover - defensive
            LOGGER.exception("Failed to load staff database")
    save_staff_database(DEFAULT_STAFF)
    LOGGER.info("已创建默认审核人员库")
    return DEFAULT_STAFF.copy()


def save_staff_database(staff_db: Dict[str, str]) -> None:
    ensure_data_dir()
    with STAFF_DB_FILE.open("w", encoding="utf-8") as handle:
        json.dump(staff_db, handle, ensure_ascii=False, indent=2)


def upsert_staff(name: str, id_last4: str) -> None:
    staff = load_staff_database()
    staff[name] = id_last4
    save_staff_database(staff)


def delete_staff(name: str) -> bool:
    staff = load_staff_database()
    if name in staff:
        del staff[name]
        save_staff_database(staff)
        return True
    return False


def match_staff_id(name: str, staff_db: Optional[Dict[str, str]] = None) -> Optional[str]:
    from .text_utils import extract_chinese_name

    if not name:
        return None
    staff = staff_db or load_staff_database()
    chinese_name = extract_chinese_name(name)
    if not chinese_name:
        return None
    return staff.get(chinese_name)


def list_staff() -> Dict[str, str]:
    staff = load_staff_database()
    return dict(sorted(staff.items(), key=lambda item: item[0]))
