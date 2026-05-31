from pathlib import Path

from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QFont
from PyQt6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTextEdit,
    QFileDialog,
    QFrame,
    QProgressBar,
    QSizePolicy,
    QMessageBox,
)
from loguru import logger

from ..core.processor import process_file, ProcessingResult


class WorkerThread(QThread):
    """
    Поток фоновой обработки файла.
    """

    finished = pyqtSignal(object)

    def __init__(self, input_path: Path, output_path: Path | None = None) -> None:
        """
        Инициализация параметров обработки.
        """

        super().__init__()
        self.input_path = input_path
        self.output_path = output_path

    def run(self) -> None:
        """
        Запуск обработки файла в отдельном потоке.
        """

        result = process_file(self.input_path, self.output_path)
        self.finished.emit(result)


class DropZone(QFrame):
    """
    Область перетаскивания файлов.
    """

    file_dropped = pyqtSignal(Path)

    _DEFAULT_STYLE = """
        DropZone {
            border: 2px dashed #5B8DB8;
            border-radius: 10px;
            background-color: #F4F8FC;
        }
        DropZone:hover {
            border-color: #2E75B6;
            background-color: #E8F1FA;
        }
    """
    _ACTIVE_STYLE = """
        DropZone {
            border: 2px solid #2E75B6;
            border-radius: 10px;
            background-color: #D6E8F7;
        }
    """

    def __init__(self, parent: QWidget | None = None) -> None:
        """
        Инициализация области перетаскивания.
        """

        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setMinimumHeight(120)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.setStyleSheet(self._DEFAULT_STYLE)

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        icon_label = QLabel("📂")
        icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon_label.setFont(QFont("Segoe UI", 28))

        self.hint_label = QLabel("Перетащите .xlsm файл сюда\nили нажмите «Выбрать файл»")
        self.hint_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.hint_label.setFont(QFont("Segoe UI", 10))
        self.hint_label.setStyleSheet("color: #5B6B7A;")

        layout.addWidget(icon_label)
        layout.addWidget(self.hint_label)

    def set_file(self, path: Path) -> None:
        """
        Отображение выбранного файла.
        """

        self.hint_label.setText(f"✓  {path.name}")
        self.hint_label.setStyleSheet("color: #1A6B2A; font-weight: bold;")

    def reset(self) -> None:
        """
        Отображение выбранного файла.
        """

        self.hint_label.setText("Перетащите .xlsm файл сюда\nили нажмите «Выбрать файл»")
        self.hint_label.setStyleSheet("color: #5B6B7A; font-weight: normal;")
        self.setStyleSheet(self._DEFAULT_STYLE)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        """
        Обработка входа файла в область перетаскивания.
        """

        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if any(u.toLocalFile().lower().endswith((".xlsm", ".xlsx")) for u in urls):
                event.acceptProposedAction()
                self.setStyleSheet(self._ACTIVE_STYLE)
                return
        event.ignore()

    def dragLeaveEvent(self, event) -> None:
        """
        Обработка выхода файла из области перетаскивания.
        """

        self.setStyleSheet(self._DEFAULT_STYLE)

    def dropEvent(self, event: QDropEvent) -> None:
        """
        Обработка сброса файла в область перетаскивания.
        """

        self.setStyleSheet(self._DEFAULT_STYLE)
        for url in event.mimeData().urls():
            path = Path(url.toLocalFile())
            if path.suffix.lower() in (".xlsm", ".xlsx"):
                self.file_dropped.emit(path)
                return


class QtLogHandler:
    """
    Обработчик вывода логов в интерфейс.
    """

    def __init__(self, text_edit: QTextEdit) -> None:
        """
        Инициализация обработчика логов.
        """

        self.text_edit = text_edit

    def write(self, message: str) -> None:
        """
        Вывод сообщения в журнал интерфейса.
        """

        stripped = message.strip()
        if not stripped:
            return
        if "ERROR" in stripped or "Ошибка" in stripped:
            color = "#C0392B"
        elif "WARNING" in stripped or "WARN" in stripped:
            color = "#D35400"
        elif "✓" in stripped or "ЗАВЕРШЕНО" in stripped:
            color = "#1A6B2A"
        else:
            color = "#2C3E50"

        self.text_edit.append(
            f'<span style="color:{color}; font-family: Consolas, monospace; font-size:9pt;">{stripped}</span>'
        )
        sb = self.text_edit.verticalScrollBar()
        sb.setValue(sb.maximum())


class MainWindow(QWidget):
    """
    Главное окно приложения.
    """

    APP_TITLE = "ЯКНО Spec Generator"
    WINDOW_MIN_W = 640
    WINDOW_MIN_H = 520

    def __init__(self) -> None:
        """
        Главное окно приложения.
        """

        super().__init__()
        self._selected_file: Path | None = None
        self._worker: WorkerThread | None = None
        self._setup_ui()
        self._setup_log_handler()

    def _setup_ui(self) -> None:
        """
        Создание элементов пользовательского интерфейса.
        """

        self.setWindowTitle(self.APP_TITLE)
        self.setMinimumSize(self.WINDOW_MIN_W, self.WINDOW_MIN_H)
        self.setStyleSheet("""
            QWidget { font-family: 'Segoe UI'; font-size: 10pt; background: #FAFBFC; }
            QPushButton {
                background-color: #2E75B6;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 20px;
                font-size: 10pt;
            }
            QPushButton:hover { background-color: #1F5893; }
            QPushButton:disabled { background-color: #A0B4C8; }
            QPushButton#btn_reset {
                background-color: #E8ECF0;
                color: #3C4858;
            }
            QPushButton#btn_reset:hover { background-color: #D0D8E0; }
        """)

        root = QVBoxLayout(self)
        root.setContentsMargins(20, 20, 20, 16)
        root.setSpacing(12)

        title = QLabel(self.APP_TITLE)
        title.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        title.setStyleSheet("color: #1A2D40;")
        root.addWidget(title)

        subtitle = QLabel(
            "Генерирует спецификацию из данных первого листа .xlsm файла"
        )
        subtitle.setWordWrap(True)
        subtitle.setStyleSheet("color: #5B6B7A; font-size: 9pt;")
        root.addWidget(subtitle)

        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setStyleSheet("color: #D0D8E0;")
        root.addWidget(line)

        self.drop_zone = DropZone()
        self.drop_zone.file_dropped.connect(self._on_file_selected)
        root.addWidget(self.drop_zone)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)

        self.btn_choose = QPushButton("📁  Выбрать файл")
        self.btn_choose.clicked.connect(self._open_file_dialog)

        self.btn_run = QPushButton("▶  Сгенерировать спецификацию")
        self.btn_run.setEnabled(False)
        self.btn_run.clicked.connect(self._run_processing)

        self.btn_reset = QPushButton("✕  Сбросить")
        self.btn_reset.setObjectName("btn_reset")
        self.btn_reset.setEnabled(False)
        self.btn_reset.clicked.connect(self._reset)

        btn_row.addWidget(self.btn_choose)
        btn_row.addWidget(self.btn_run, stretch=1)
        btn_row.addWidget(self.btn_reset)
        root.addLayout(btn_row)

        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.setVisible(False)
        self.progress.setFixedHeight(6)
        self.progress.setStyleSheet("""
            QProgressBar { border: none; background: #E8ECF0; border-radius: 3px; }
            QProgressBar::chunk { background: #2E75B6; border-radius: 3px; }
        """)
        root.addWidget(self.progress)

        self.status_label = QLabel("Выберите файл для начала работы")
        self.status_label.setStyleSheet("color: #5B6B7A; font-size: 9pt;")
        root.addWidget(self.status_label)

        log_title = QLabel("Журнал")
        log_title.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
        log_title.setStyleSheet("color: #3C4858;")
        root.addWidget(log_title)

        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setMinimumHeight(180)
        self.log_view.setStyleSheet("""
            QTextEdit {
                background: #F0F3F5;
                border: 1px solid #D0D8E0;
                border-radius: 6px;
                padding: 6px;
                font-family: Consolas, monospace;
                font-size: 9pt;
                color: #2C3E50;
            }
        """)
        root.addWidget(self.log_view, stretch=1)

    def _setup_log_handler(self) -> None:
        """
        Настройка вывода логов в интерфейс.
        """

        handler = QtLogHandler(self.log_view)
        logger.add(
            handler.write,
            format="{time:HH:mm:ss} | {level:<7} | {message}",
            level="DEBUG",
            colorize=False,
        )

    def _open_file_dialog(self) -> None:
        """
        Настройка вывода логов в интерфейс.
        """

        path_str, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл Excel",
            "",
            "Excel с макросами (*.xlsm);;Excel (*.xlsx);;Все файлы (*)",
        )
        if path_str:
            self._on_file_selected(Path(path_str))

    def _on_file_selected(self, path: Path) -> None:
        """
        Обработка выбора файла.
        """

        self._selected_file = path
        self.drop_zone.set_file(path)
        self.btn_run.setEnabled(True)
        self.btn_reset.setEnabled(True)
        self._set_status(f"Выбран файл: {path.name}", "info")
        logger.info(f"Файл выбран: {path}")

    def _run_processing(self) -> None:
        """
        Запуск обработки выбранного файла.
        """

        if not self._selected_file:
            return

        self._set_busy(True)
        self._set_status("Обработка...", "info")
        self.log_view.clear()

        self._worker = WorkerThread(self._selected_file)
        self._worker.finished.connect(self._on_processing_done)
        self._worker.start()

    def _on_processing_done(self, result: ProcessingResult) -> None:
        """
        Обработка завершения генерации спецификации.
        """

        self._set_busy(False)
        self._worker = None

        if result.success:
            self._set_status(f"✓ Готово! Сохранено: {result.output_path.name}", "success")
            logger.info(result.summary())
        else:
            self._set_status(f"✗ Ошибка: {result.error}", "error")
            QMessageBox.critical(self, "Ошибка обработки", result.error)

    def _reset(self) -> None:
        """
        Сброс состояния интерфейса.
        """

        self._selected_file = None
        self.drop_zone.reset()
        self.btn_run.setEnabled(False)
        self.btn_reset.setEnabled(False)
        self.log_view.clear()
        self._set_status("Выберите файл для начала работы", "info")

    def _set_busy(self, busy: bool) -> None:
        """
        Переключение режима занятости интерфейса.
        """

        self.progress.setVisible(busy)
        self.btn_run.setEnabled(not busy)
        self.btn_choose.setEnabled(not busy)
        self.btn_reset.setEnabled(not busy)

    def _set_status(self, text: str, kind: str = "info") -> None:
        """
        Обновление строки состояния.
        """

        colors = {
            "info": "#5B6B7A",
            "success": "#1A6B2A",
            "error": "#C0392B",
        }
        color = colors.get(kind, colors["info"])
        self.status_label.setText(text)
        self.status_label.setStyleSheet(f"color: {color}; font-size: 9pt;")
