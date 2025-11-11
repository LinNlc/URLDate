from __future__ import annotations

from typing import Any, Callable, Optional

from PySide6.QtCore import QObject, QThread, Signal


ProgressHook = Callable[[int, int, str], None]
MessageHook = Callable[[str, str], None]
TaskCallable = Callable[[ProgressHook, MessageHook], Any]


class Worker(QObject):
    finished = Signal(object)
    error = Signal(str)
    progress = Signal(int, int, str)
    message = Signal(str, str)

    def __init__(self, task: TaskCallable, *args, **kwargs) -> None:
        super().__init__()
        self._task = task
        self._args = args
        self._kwargs = kwargs

    def run(self) -> None:
        try:
            result = self._task(self._emit_progress, self._emit_message, *self._args, **self._kwargs)
            self.finished.emit(result)
        except Exception as exc:  # pragma: no cover - executed in thread
            self.error.emit(str(exc))

    def _emit_progress(self, current: int, total: int, description: str) -> None:
        self.progress.emit(current, total, description)

    def _emit_message(self, message: str, level: str = "info") -> None:
        self.message.emit(message, level)


class WorkerController:
    def __init__(self) -> None:
        self._thread: Optional[QThread] = None
        self._worker: Optional[Worker] = None

    def start(self, worker: Worker) -> None:
        self.stop()
        thread = QThread()
        self._thread = thread
        self._worker = worker
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.finished.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        worker.error.connect(thread.quit)
        worker.error.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        thread.start()

    def stop(self) -> None:
        if self._thread and self._thread.isRunning():
            self._thread.quit()
            self._thread.wait()
        self._thread = None
        self._worker = None
