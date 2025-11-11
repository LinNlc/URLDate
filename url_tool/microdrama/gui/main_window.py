from __future__ import annotations

from pathlib import Path
from typing import List

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QProgressBar,
    QSizePolicy,
    QSpacerItem,
    QVBoxLayout,
    QWidget,
)

from ..core.config_store import ensure_mode, load_mode
from ..core.excel_processor import ExcelProcessor
from ..core.version_checker import build_update_url, fetch_latest_version
from ..utils.logger import attach_dispatcher, get_logger, register_callback
from ..utils.workers import Worker, WorkerController
from .staff_dialog import StaffDialog

APP_TITLE = "微剧URL转换工具 v1.7"


class MainWindow(QMainWindow):
    log_signal = Signal(str, str)

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(900, 600)
        self._logger = get_logger(__name__)
        attach_dispatcher()
        register_callback(lambda message, level: self.log_signal.emit(message, level))
        self.log_signal.connect(self._append_log)
        self._worker_controller = WorkerController()
        self._selected_file: Path | None = None
        self._setup_ui()
        self._load_mode()

    # ------------------------------------------------------------------
    # UI setup
    # ------------------------------------------------------------------
    def _setup_ui(self) -> None:
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        # File selection area
        file_group = QGroupBox("选择Excel文件")
        file_layout = QGridLayout()
        file_group.setLayout(file_layout)

        self.file_path_edit = QLineEdit()
        self.file_path_edit.setReadOnly(True)
        file_button = QPushButton("浏览...")
        file_button.clicked.connect(self._select_file)

        file_layout.addWidget(QLabel("文件路径"), 0, 0)
        file_layout.addWidget(self.file_path_edit, 0, 1)
        file_layout.addWidget(file_button, 0, 2)

        layout.addWidget(file_group)

        # Mode selection
        mode_group = QGroupBox("处理模式")
        mode_layout = QHBoxLayout()
        mode_group.setLayout(mode_layout)

        self.mode_combo = QComboBox()
        self.mode_combo.addItem("模式1：每50条拆分", 1)
        self.mode_combo.addItem("模式2：整表处理", 2)
        self.mode_combo.currentIndexChanged.connect(self._mode_changed)

        mode_layout.addWidget(QLabel("当前模式"))
        mode_layout.addWidget(self.mode_combo)
        mode_layout.addItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))

        layout.addWidget(mode_group)

        # Action buttons
        button_bar = QHBoxLayout()
        self.process_button = QPushButton("开始处理")
        self.process_button.clicked.connect(self._start_processing)
        self.staff_button = QPushButton("人员库管理")
        self.staff_button.clicked.connect(self._open_staff_dialog)
        self.update_button = QPushButton("检查更新")
        self.update_button.clicked.connect(self._check_update)

        button_bar.addWidget(self.process_button)
        button_bar.addWidget(self.staff_button)
        button_bar.addWidget(self.update_button)
        button_bar.addItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))

        layout.addLayout(button_bar)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        layout.addWidget(self.progress_bar)

        # Log output
        log_group = QGroupBox("日志输出")
        log_layout = QVBoxLayout()
        log_group.setLayout(log_layout)

        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        log_layout.addWidget(self.log_view)

        layout.addWidget(log_group, stretch=1)

    # ------------------------------------------------------------------
    # Mode handling
    # ------------------------------------------------------------------
    def _load_mode(self) -> None:
        mode = load_mode()
        index = self.mode_combo.findData(mode)
        if index != -1:
            self.mode_combo.setCurrentIndex(index)

    def _mode_changed(self) -> None:
        mode = self.mode_combo.currentData()
        ensure_mode(mode)
        self._append_log(f"已切换到模式{mode}", "info")

    # ------------------------------------------------------------------
    # File selection
    # ------------------------------------------------------------------
    def _select_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            self._selected_file = Path(file_path)
            self.file_path_edit.setText(str(self._selected_file))
            self._append_log(f"已选择文件: {self._selected_file}", "info")

    # ------------------------------------------------------------------
    # Processing logic
    # ------------------------------------------------------------------
    def _start_processing(self) -> None:
        if not self._selected_file or not self._selected_file.exists():
            QMessageBox.warning(self, APP_TITLE, "请先选择有效的Excel文件")
            return

        mode = self.mode_combo.currentData()
        self._set_running_state(True)
        self.progress_bar.setValue(0)
        self.log_view.appendPlainText("\n==== 开始处理 ====")

        def task(progress_cb, message_cb):
            processor = ExcelProcessor(
                logger_callback=message_cb,
                progress_callback=progress_cb,
            )
            return processor.process(str(self._selected_file), mode)

        worker = Worker(task)
        worker.progress.connect(self._on_progress)
        worker.message.connect(self._append_log)
        worker.finished.connect(self._on_finished)
        worker.error.connect(self._on_error)
        self._worker_controller.start(worker)

    def _set_running_state(self, running: bool) -> None:
        self.process_button.setEnabled(not running)
        self.staff_button.setEnabled(not running)
        self.update_button.setEnabled(not running)

    def _on_progress(self, current: int, total: int, description: str) -> None:
        if total <= 0:
            return
        value = int(current / total * 100)
        self.progress_bar.setValue(value)
        self.progress_bar.setFormat(f"{description} {current}/{total}")

    def _on_finished(self, result: List[Path]) -> None:
        self._set_running_state(False)
        self.progress_bar.setValue(100)
        if result:
            files = "\n".join(str(path) for path in result)
            QMessageBox.information(self, APP_TITLE, f"处理完成!\n输出文件:\n{files}")
        self._append_log("==== 处理完成 ====", "success")

    def _on_error(self, message: str) -> None:
        self._set_running_state(False)
        QMessageBox.critical(self, APP_TITLE, f"处理失败: {message}")
        self._append_log(message, "error")

    # ------------------------------------------------------------------
    # Staff dialog
    # ------------------------------------------------------------------
    def _open_staff_dialog(self) -> None:
        dialog = StaffDialog(self)
        dialog.exec()

    # ------------------------------------------------------------------
    # Update check
    # ------------------------------------------------------------------
    def _check_update(self) -> None:
        latest = fetch_latest_version()
        if not latest:
            QMessageBox.warning(self, APP_TITLE, "无法获取最新版本信息")
            return
        url = build_update_url(latest)
        QMessageBox.information(self, APP_TITLE, f"最新版本: {latest}\n下载地址:\n{url}")

    # ------------------------------------------------------------------
    # Logging
    # ------------------------------------------------------------------
    def _append_log(self, message: str, level: str = "info") -> None:
        self.log_view.appendPlainText(f"[{level.upper()}] {message}")
        cursor = self.log_view.textCursor()
        cursor.movePosition(cursor.End)
        self.log_view.setTextCursor(cursor)


def run_app() -> None:
    app = QApplication.instance() or QApplication([])
    window = MainWindow()
    window.show()
    app.exec()

