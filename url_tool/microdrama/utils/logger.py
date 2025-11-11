from __future__ import annotations

import logging
from typing import Callable, List

_LOGGER_NAME = "microdrama"
_callbacks: List[Callable[[str, str], None]] = []


def configure_logging(level: int = logging.INFO) -> None:
    logger = logging.getLogger(_LOGGER_NAME)
    if logger.handlers:
        return
    logger.setLevel(level)
    handler = logging.StreamHandler()
    formatter = logging.Formatter("[%(asctime)s] %(levelname)s - %(message)s", "%H:%M:%S")
    handler.setFormatter(formatter)
    logger.addHandler(handler)


def get_logger(name: str | None = None) -> logging.Logger:
    configure_logging()
    if name:
        return logging.getLogger(f"{_LOGGER_NAME}.{name}")
    return logging.getLogger(_LOGGER_NAME)


class UILogHandler(logging.Handler):
    def __init__(self, callback: Callable[[str, str], None]) -> None:
        super().__init__()
        self.callback = callback

    def emit(self, record: logging.LogRecord) -> None:  # pragma: no cover - passthrough
        message = self.format(record)
        level = record.levelname.lower()
        self.callback(message, level)


def register_callback(callback: Callable[[str, str], None]) -> None:
    _callbacks.append(callback)


class CallbackDispatcher(logging.Handler):
    def emit(self, record: logging.LogRecord) -> None:  # pragma: no cover - simple pass-through
        message = record.getMessage()
        level = record.levelname.lower()
        for callback in list(_callbacks):
            callback(message, level)


def attach_dispatcher() -> None:
    logger = get_logger()
    for handler in logger.handlers:
        if isinstance(handler, CallbackDispatcher):
            return
    dispatcher = CallbackDispatcher()
    dispatcher.setLevel(logging.INFO)
    logger.addHandler(dispatcher)
