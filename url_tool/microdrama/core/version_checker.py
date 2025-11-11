from __future__ import annotations

from typing import Optional

import requests

from ..utils.logger import get_logger

LOGGER = get_logger(__name__)

VERSION_CHECK_URL = "https://raw.githubusercontent.com/LinNlc/-URL-/main/version.json"
UPDATE_DOWNLOAD_URL = "https://github.com/LinNlc/-URL-/releases/download/v{version}/微剧URL转换工具v{version}.exe"


def fetch_latest_version() -> Optional[str]:
    try:
        response = requests.get(VERSION_CHECK_URL, timeout=5)
        response.raise_for_status()
        data = response.json()
        return data.get("version")
    except Exception as exc:  # pragma: no cover - network call
        LOGGER.warning("获取版本信息失败: %s", exc)
        return None


def build_update_url(version: str) -> str:
    return UPDATE_DOWNLOAD_URL.format(version=version)
