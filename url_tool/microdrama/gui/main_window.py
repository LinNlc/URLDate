from __future__ import annotations

from pathlib import Path
from typing import List

from PySide6.QtCore import (
    QEasingCurve,
    QEvent,
    QPropertyAnimation,
    Qt,
    QTimer,
    Signal,
    QAbstractAnimation,
)
from PySide6.QtGui import QColor, QFont
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
    QGraphicsDropShadowEffect,
    QGraphicsOpacityEffect,
    QGraphicsColorizeEffect,
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
        # 修改开始: 注入 macOS 风格主题与动画所需的初始化
        self._button_effects: dict[QPushButton, QGraphicsDropShadowEffect] = {}
        self._button_animations: dict[QPushButton, QPropertyAnimation] = {}
        self._intro_animation: QPropertyAnimation | None = None
        self._log_highlight_effect: QGraphicsColorizeEffect | None = None
        self._log_pulse_anim: QPropertyAnimation | None = None
        self._setup_ui()
        self._apply_macos_theme()
        self._setup_animations()
        # 修改结束
        self._load_mode()

    # ------------------------------------------------------------------
    # UI setup
    # ------------------------------------------------------------------
    def _setup_ui(self) -> None:
        central_widget = QWidget(self)
        central_widget.setObjectName("centralWidget")
        self._central_widget = central_widget
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
        self.file_path_edit.setPlaceholderText("请选择 Excel 文件")
        file_button = QPushButton("浏览...")
        file_button.clicked.connect(self._select_file)
        file_button.setCursor(Qt.PointingHandCursor)
        self._file_button = file_button

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
        self.mode_combo.setCursor(Qt.PointingHandCursor)

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

        for button in (self.process_button, self.staff_button, self.update_button):
            button.setCursor(Qt.PointingHandCursor)

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
        self.log_view.setObjectName("logView")
        self.log_view.setReadOnly(True)
        log_layout.addWidget(self.log_view)

        layout.addWidget(log_group, stretch=1)

        # 修改开始: 记录需要美化的组件列表，供后续动画与风格调用
        self._group_boxes = [file_group, mode_group, log_group]
        self._all_buttons = [file_button, self.process_button, self.staff_button, self.update_button]
        # 修改结束

    # ------------------------------------------------------------------
    # macOS 风格主题与动画配置
    # ------------------------------------------------------------------
    def _apply_macos_theme(self) -> None:
        # 修改开始: 应用苹果风格的配色、圆角与阴影
        self.setFont(QFont("SF Pro Text", 11))
        accent_color = "#0A84FF"
        stylesheet = f"""
            QWidget#centralWidget {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 rgba(238, 242, 255, 1),
                    stop:1 rgba(216, 228, 255, 1));
            }}
            QGroupBox {{
                background-color: rgba(255, 255, 255, 230);
                border: 1px solid rgba(255, 255, 255, 200);
                border-radius: 18px;
                margin-top: 18px;
                padding: 22px 20px 18px 20px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 0 10px 0;
                color: rgba(44, 46, 51, 0.88);
                font-size: 16px;
                font-weight: 600;
            }}
            QLabel {{
                color: rgba(44, 46, 51, 0.88);
                font-size: 14px;
            }}
            QLineEdit {{
                border: 1px solid rgba(180, 185, 200, 0.7);
                border-radius: 12px;
                padding: 8px 12px;
                background-color: rgba(255, 255, 255, 0.92);
                selection-background-color: {accent_color};
            }}
            QComboBox {{
                border: 1px solid rgba(180, 185, 200, 0.7);
                border-radius: 12px;
                padding: 6px 30px 6px 12px;
                background-color: rgba(255, 255, 255, 0.92);
            }}
            QComboBox::drop-down {{
                width: 28px;
                border: none;
            }}
            QComboBox QAbstractItemView {{
                border-radius: 12px;
                selection-background-color: {accent_color};
                padding: 6px;
            }}
            QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {accent_color}, stop:1 #055CBA);
                border-radius: 14px;
                color: white;
                font-size: 15px;
                font-weight: 600;
                padding: 10px 22px;
                border: none;
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #2EA6FF, stop:1 #0A84FF);
            }}
            QPushButton:pressed {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #055CBA, stop:1 #034A99);
            }}
            QProgressBar {{
                border: 1px solid rgba(180, 185, 200, 0.6);
                border-radius: 12px;
                background-color: rgba(255, 255, 255, 0.6);
                padding: 3px;
                text-align: center;
                font-weight: 600;
            }}
            QProgressBar::chunk {{
                border-radius: 10px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 {accent_color}, stop:1 #74B7FF);
            }}
            QPlainTextEdit#logView {{
                border: none;
                border-radius: 16px;
                padding: 12px;
                background-color: rgba(246, 248, 255, 0.9);
                font-family: "SF Mono", "Menlo", monospace;
                font-size: 13px;
            }}
        """
        self.setStyleSheet(stylesheet)

        for group in self._group_boxes:
            shadow = QGraphicsDropShadowEffect(self)
            shadow.setBlurRadius(28)
            shadow.setOffset(0, 14)
            shadow.setColor(QColor(15, 24, 44, 60))
            group.setGraphicsEffect(shadow)

        self._log_highlight_effect = QGraphicsColorizeEffect(self.log_view)
        self._log_highlight_effect.setColor(QColor(10, 132, 255))
        self._log_highlight_effect.setStrength(0.0)
        self.log_view.setGraphicsEffect(self._log_highlight_effect)
        # 修改结束

    def _setup_animations(self) -> None:
        # 修改开始: 构建全局淡入动画和按钮悬停动效
        opacity_effect = QGraphicsOpacityEffect(self._central_widget)
        self._central_widget.setGraphicsEffect(opacity_effect)
        opacity_effect.setOpacity(0.0)

        self._intro_animation = QPropertyAnimation(opacity_effect, b"opacity", self)
        self._intro_animation.setStartValue(0.0)
        self._intro_animation.setEndValue(1.0)
        self._intro_animation.setDuration(800)
        self._intro_animation.setEasingCurve(QEasingCurve.OutCubic)
        self._intro_animation.start()
        self._intro_animation.finished.connect(lambda: self._central_widget.setGraphicsEffect(None))

        for button in self._all_buttons:
            effect = QGraphicsDropShadowEffect(self)
            effect.setBlurRadius(16)
            effect.setOffset(0, 6)
            effect.setColor(QColor(10, 132, 255, 100))
            button.setGraphicsEffect(effect)

            animation = QPropertyAnimation(effect, b"blurRadius", self)
            animation.setStartValue(16)
            animation.setEndValue(32)
            animation.setDuration(250)
            animation.setEasingCurve(QEasingCurve.OutCubic)
            animation.setDirection(QAbstractAnimation.Backward)

            self._button_effects[button] = effect
            self._button_animations[button] = animation
            button.installEventFilter(self)

        if self._log_highlight_effect:
            self._log_pulse_anim = QPropertyAnimation(self._log_highlight_effect, b"strength", self)
            self._log_pulse_anim.setStartValue(0.0)
            self._log_pulse_anim.setEndValue(0.6)
            self._log_pulse_anim.setDuration(320)
            self._log_pulse_anim.setEasingCurve(QEasingCurve.OutCubic)
            self._log_pulse_anim.setDirection(QAbstractAnimation.Backward)
        # 修改结束

    def eventFilter(self, obj, event):  # type: ignore[override]
        if obj in self._button_animations:
            effect = self._button_effects[obj]
            animation = self._button_animations[obj]
            if event.type() == QEvent.Enter:
                effect.setColor(QColor(10, 132, 255, 150))
                effect.setYOffset(12)
                animation.setDirection(QAbstractAnimation.Forward)
                animation.start()
            elif event.type() == QEvent.Leave:
                effect.setColor(QColor(10, 132, 255, 100))
                effect.setYOffset(6)
                animation.setDirection(QAbstractAnimation.Backward)
                animation.start()
            elif event.type() == QEvent.MouseButtonPress:
                effect.setYOffset(3)
            elif event.type() == QEvent.MouseButtonRelease:
                effect.setYOffset(12 if obj.underMouse() else 6)
            return False
        return super().eventFilter(obj, event)

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
        if self._log_pulse_anim and self._log_highlight_effect:
            self._log_pulse_anim.stop()
            self._log_highlight_effect.setStrength(0.0)
            self._log_pulse_anim.setDirection(QAbstractAnimation.Forward)
            self._log_pulse_anim.start()

            def _fade_back() -> None:
                if self._log_pulse_anim:
                    self._log_pulse_anim.setDirection(QAbstractAnimation.Backward)
                    self._log_pulse_anim.start()

            QTimer.singleShot(260, _fade_back)


def run_app() -> None:
    app = QApplication.instance() or QApplication([])
    window = MainWindow()
    window.show()
    app.exec()

