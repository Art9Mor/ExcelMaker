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
    QCheckBox,
    QGroupBox,
)
from loguru import logger

from ..core.models import SpecDocument
from ..core.parser import load_and_parse
from ..core.processor import ProcessingResult
from ..core.writer import write_spec_to_sheet
from ..core.word_generator import save_as_word


class WorkerThread(QThread):
    finished = pyqtSignal(object)

    def __init__(
            self,
            wb,
            doc: SpecDocument,
            output_excel_path: Path,
            output_word_path: Path | None = None,
    ) -> None:
        super().__init__()
        self.wb = wb
        self.doc = doc
        self.output_excel_path = output_excel_path
        self.output_word_path = output_word_path

    def run(self) -> None:
        try:
            # Сохраняем Excel
            write_spec_to_sheet(self.wb, self.doc)
            self.wb.save(str(self.output_excel_path))

            # Сохраняем Word если нужно
            if self.output_word_path:
                save_as_word(self.doc, self.output_word_path)

            result = ProcessingResult(
                success=True,
                doc=self.doc,
                output_path=self.output_excel_path,
                word_path=self.output_word_path,
            )
        except Exception as e:
            logger.exception("Ошибка записи")
            result = ProcessingResult(success=False, error=str(e))

        self.finished.emit(result)

    def stop(self):
        """Безопасная остановка потока"""
        self.quit()
        self.wait(5000)  # Ждём до 5 секунд


class DropZone(QFrame):
    file_dropped = pyqtSignal(Path)

    _STYLE = """
        QFrame {
            border: 2px dashed #5B8DB8;
            border-radius: 10px;
            background-color: #F4F8FC;
        }
    """

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setMinimumHeight(120)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.setStyleSheet(self._STYLE)

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
        self.hint_label.setText(f"✓  {path.name}")
        self.hint_label.setStyleSheet("color: #1A6B2A; font-weight: bold;")

    def reset(self) -> None:
        self.hint_label.setText("Перетащите .xlsm файл сюда\nили нажмите «Выбрать файл»")
        self.hint_label.setStyleSheet("color: #5B6B7A; font-weight: normal;")

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if any(u.toLocalFile().lower().endswith((".xlsm", ".xlsx")) for u in urls):
                event.acceptProposedAction()
                return
        event.ignore()

    def dragLeaveEvent(self, event) -> None:
        pass

    def dropEvent(self, event: QDropEvent) -> None:
        for url in event.mimeData().urls():
            path = Path(url.toLocalFile())
            if path.suffix.lower() in (".xlsm", ".xlsx"):
                self.file_dropped.emit(path)
                return


class QtLogHandler:
    def __init__(self, text_edit: QTextEdit) -> None:
        self.text_edit = text_edit

    def write(self, message: str) -> None:
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
    APP_TITLE = "ЯКНО Spec Generator"
    WINDOW_MIN_W = 640
    WINDOW_MIN_H = 580

    def __init__(self) -> None:
        super().__init__()
        self._selected_file: Path | None = None
        self._output_folder: Path | None = None
        self._worker: WorkerThread | None = None
        self._setup_ui()
        self._setup_log_handler()

    def _setup_ui(self) -> None:
        self.setWindowTitle(self.APP_TITLE)
        self.setMinimumSize(self.WINDOW_MIN_W, self.WINDOW_MIN_H)
        self.setStyleSheet("""
            QWidget { font-family: 'Segoe UI'; font-size: 10pt; background: #FAFBFC; color: #2C3E50; }
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
            QPushButton#btn_output {
                background-color: #E8ECF0;
                color: #3C4858;
                padding: 6px 12px;
                font-size: 9pt;
            }
            QPushButton#btn_output:hover { background-color: #D0D8E0; }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #D0D8E0;
                border-radius: 6px;
                margin-top: 8px;
                padding-top: 8px;
                background-color: white;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }
            QCheckBox {
                spacing: 5px;
                color: #2C3E50;
            }
            QCheckBox::indicator {
                width: 16px;
                height: 16px;
            }
            QLabel {
                color: #2C3E50;
            }
            QProgressBar {
                border: none;
                background: #E8ECF0;
                border-radius: 3px;
                height: 6px;
            }
            QProgressBar::chunk {
                background: #2E75B6;
                border-radius: 3px;
            }
            QTextEdit {
                background: #F8F9FA;
                border: 1px solid #D0D8E0;
                border-radius: 6px;
                padding: 6px;
                font-family: Consolas, monospace;
                font-size: 9pt;
                color: #2C3E50;
            }
        """)

        root = QVBoxLayout(self)
        root.setContentsMargins(20, 20, 20, 16)
        root.setSpacing(12)

        title = QLabel(self.APP_TITLE)
        title.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        title.setStyleSheet("color: #1A2D40;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        root.addWidget(title)

        subtitle = QLabel("Генерирует спецификацию из данных первого листа .xlsm файла")
        subtitle.setWordWrap(True)
        subtitle.setStyleSheet("color: #5B6B7A; font-size: 9pt;")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        root.addWidget(subtitle)

        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setStyleSheet("background-color: #D0D8E0; max-height: 1px;")
        root.addWidget(line)

        self.drop_zone = DropZone()
        self.drop_zone.file_dropped.connect(self._on_file_selected)
        root.addWidget(self.drop_zone)

        format_group = QGroupBox("Формат вывода")
        format_group.setStyleSheet("QGroupBox { background-color: white; }")
        format_layout = QHBoxLayout(format_group)

        self.check_excel = QCheckBox("Excel (XLSX)")
        self.check_excel.setChecked(True)
        self.check_word = QCheckBox("Word (DOCX)")

        format_layout.addWidget(self.check_excel)
        format_layout.addWidget(self.check_word)
        format_layout.addStretch()
        root.addWidget(format_group)

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

        output_row = QHBoxLayout()
        output_row.setSpacing(8)

        self.output_path_label = QLabel("📁 Результат будет сохранён рядом с исходным файлом")
        self.output_path_label.setStyleSheet("color: #5B6B7A; font-size: 9pt;")

        self.btn_choose_output = QPushButton("Выбрать папку сохранения")
        self.btn_choose_output.setObjectName("btn_output")
        self.btn_choose_output.setEnabled(False)
        self.btn_choose_output.clicked.connect(self._choose_output_folder)

        output_row.addWidget(self.output_path_label, stretch=1)
        output_row.addWidget(self.btn_choose_output)
        root.addLayout(output_row)

        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.setVisible(False)
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
        root.addWidget(self.log_view, stretch=1)

    def _setup_log_handler(self) -> None:
        handler = QtLogHandler(self.log_view)
        logger.add(
            handler.write,
            format="{time:HH:mm:ss} | {level:<7} | {message}",
            level="DEBUG",
            colorize=False,
        )

    def _open_file_dialog(self) -> None:
        path_str, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл Excel",
            "",
            "Excel с макросами (*.xlsm);;Excel (*.xlsx);;Все файлы (*)",
        )
        if path_str:
            self._on_file_selected(Path(path_str))

    def _choose_output_folder(self) -> None:
        initial_dir = str(self._selected_file.parent) if self._selected_file else ""
        folder = QFileDialog.getExistingDirectory(
            self,
            "Выберите папку для сохранения результата",
            initial_dir,
        )
        if folder:
            self._output_folder = Path(folder)
            self.output_path_label.setText(f"📁 Результат будет сохранён в: {folder}")
            logger.info(f"Выбрана папка сохранения: {folder}")
        else:
            self._output_folder = None
            self.output_path_label.setText("📁 Результат будет сохранён рядом с исходным файлом")

    def _on_file_selected(self, path: Path) -> None:
        self._selected_file = path
        self._output_folder = None
        self.output_path_label.setText("📁 Результат будет сохранён рядом с исходным файлом")
        self.drop_zone.set_file(path)
        self.btn_run.setEnabled(True)
        self.btn_reset.setEnabled(True)
        self.btn_choose_output.setEnabled(True)
        self._set_status(f"Выбран файл: {path.name}", "info")
        logger.info(f"Файл выбран: {path}")

    def _run_processing(self) -> None:
        if not self._selected_file:
            return

        if not self.check_excel.isChecked() and not self.check_word.isChecked():
            QMessageBox.warning(self, "Выбор формата", "Выберите хотя бы один формат вывода")
            return

        # Формируем путь для СОХРАНЕНИЯ (не перезаписываем исходный)
        if self._output_folder:
            base_path = self._output_folder / f"{self._selected_file.stem}_specification"
        else:
            base_path = self._selected_file.parent / f"{self._selected_file.stem}_specification"

        output_excel_path = None
        output_word_path = None

        if self.check_excel.isChecked():
            output_excel_path = base_path.with_suffix(".xlsm")
            logger.info(f"Будет сохранён Excel: {output_excel_path}")

        if self.check_word.isChecked():
            output_word_path = base_path.with_suffix(".docx")
            logger.info(f"Будет сохранён Word: {output_word_path}")

        self._set_busy(True)
        self._set_status("Обработка...", "info")
        self.log_view.clear()

        try:
            wb, doc = load_and_parse(self._selected_file)
        except Exception as e:
            logger.exception("Ошибка загрузки файла")
            self._set_busy(False)
            self._set_status(f"✗ Ошибка: {e}", "error")
            QMessageBox.critical(self, "Ошибка загрузки", str(e))
            return

        # Останавливаем предыдущий поток если есть
        if self._worker is not None:
            self._worker.stop()
            self._worker = None

        self._worker = WorkerThread(wb, doc, output_excel_path, output_word_path)
        self._worker.finished.connect(self._on_processing_done)
        self._worker.start()

    def _on_processing_done(self, result: ProcessingResult) -> None:
        self._set_busy(False)

        if result.success:
            msg = "✅ Файлы успешно сохранены:\n\n"
            if self.check_excel.isChecked() and result.output_path:
                msg += f"📊 Excel: {result.output_path}\n"
            if self.check_word.isChecked():
                if result.word_path:
                    msg += f"📄 Word: {result.word_path}"
                elif result.output_path:
                    word_path = result.output_path.with_suffix(".docx")
                    if word_path.exists():
                        msg += f"📄 Word: {word_path}"

            self._set_status("✓ Готово!", "success")
            logger.info(result.summary())
            QMessageBox.information(self, "Генерация завершена", msg)
        else:
            self._set_status(f"✗ Ошибка: {result.error}", "error")
            QMessageBox.critical(self, "Ошибка обработки", result.error)

        # Очищаем worker после завершения
        if self._worker is not None:
            self._worker.deleteLater()
            self._worker = None

    def _reset(self) -> None:
        # Останавливаем worker если есть
        if self._worker is not None:
            self._worker.stop()
            self._worker = None

        self._selected_file = None
        self._output_folder = None
        self.drop_zone.reset()
        self.output_path_label.setText("📁 Результат будет сохранён рядом с исходным файлом")
        self.btn_run.setEnabled(False)
        self.btn_reset.setEnabled(False)
        self.btn_choose_output.setEnabled(False)
        self.log_view.clear()
        self._set_status("Выберите файл для начала работы", "info")

    def _set_busy(self, busy: bool) -> None:
        self.progress.setVisible(busy)
        self.btn_run.setEnabled(not busy)
        self.btn_choose.setEnabled(not busy)
        self.btn_reset.setEnabled(not busy)
        self.btn_choose_output.setEnabled(not busy and self._selected_file is not None)

    def _set_status(self, text: str, kind: str = "info") -> None:
        colors = {
            "info": "#5B6B7A",
            "success": "#1A6B2A",
            "error": "#C0392B",
        }
        color = colors.get(kind, colors["info"])
        self.status_label.setText(text)
        self.status_label.setStyleSheet(f"color: {color}; font-size: 9pt;")

    def closeEvent(self, event):
        """Обработка закрытия окна"""
        if self._worker is not None:
            self._worker.stop()
            self._worker = None
        event.accept()
