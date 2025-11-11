from __future__ import annotations

import concurrent.futures
import io
import time
from pathlib import Path
from typing import Dict, Iterable, Optional, Tuple

import requests
from PIL import Image

from ..utils.logger import get_logger

LOGGER = get_logger(__name__)


HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
}


def download_image(url: str, timeout: int = 15) -> Optional[bytes]:
    try:
        response = requests.get(url, headers=HEADERS, timeout=timeout)
        response.raise_for_status()
        content_type = response.headers.get("content-type", "").lower()
        if not content_type.startswith("image/"):
            return None
        return response.content
    except Exception as exc:
        LOGGER.warning("下载图片失败 %s: %s", url, exc)
        return None


def download_images_concurrently(urls: Iterable[str], max_workers: int = 8) -> Dict[int, Optional[bytes]]:
    tasks = list(enumerate(urls))
    results: Dict[int, Optional[bytes]] = {}

    def _download(index_url: Tuple[int, str]) -> Tuple[int, Optional[bytes]]:
        idx, link = index_url
        return idx, download_image(link)

    if not tasks:
        return results

    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        for idx, image_data in executor.map(_download, tasks):
            results[idx] = image_data
    return results


def resize_image(image_data: bytes, target_width: int, target_height: int) -> Optional[bytes]:
    try:
        pil_image = Image.open(io.BytesIO(image_data))
        pil_image = pil_image.convert("RGB")
        try:
            resample_method = Image.Resampling.LANCZOS  # Pillow >= 10
        except AttributeError:  # pragma: no cover - fallback for old Pillow
            resample_method = getattr(Image, "LANCZOS", getattr(Image, "BICUBIC", 3))
        pil_image = pil_image.resize((target_width, target_height), resample_method)
        buffer = io.BytesIO()
        pil_image.save(buffer, format="JPEG", quality=80)
        return buffer.getvalue()
    except Exception as exc:  # pragma: no cover - defensive
        LOGGER.warning("处理图片失败: %s", exc)
        return None


def save_temp_image(image_bytes: bytes, identifier: str, directory: Path) -> Path:
    directory.mkdir(parents=True, exist_ok=True)
    filename = f"temp_image_{identifier}_{int(time.time())}.jpg"
    path = directory / filename
    with path.open("wb") as handle:
        handle.write(image_bytes)
    return path
