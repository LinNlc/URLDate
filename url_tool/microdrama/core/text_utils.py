from __future__ import annotations

import re
from typing import Optional


def extract_chinese_name(name: Optional[str]) -> str:
    if not name:
        return ""
    chinese_chars = re.findall(r"[\u4e00-\u9fff]+", str(name))
    return "".join(chinese_chars)


def check_content_length(text: Optional[str]) -> int:
    if not text:
        return 0
    return len(str(text).strip())


def has_invalid_actor_delimiter(name: Optional[str]) -> bool:
    if not name:
        return False
    return "-" in str(name)


def is_valid_url(url: Optional[str]) -> bool:
    if not url or not isinstance(url, str):
        return False
    url_pattern = re.compile(
        r"^https?://"
        r"(?:(?:[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?\.)+[A-Z]{2,6}\.?|"
        r"localhost|"
        r"\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})"
        r"(?::\d+)?"
        r"(?:/?|[/?]\S+)$",
        re.IGNORECASE,
    )
    return bool(url_pattern.match(url.strip()))
