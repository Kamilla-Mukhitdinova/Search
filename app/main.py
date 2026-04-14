import os
import string
import subprocess
import sys
import zipfile
import xml.etree.ElementTree as ET

from PySide6.QtCore import QObject, QThread, Qt, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)


class SearchWorker(QObject):
    result_found = Signal(str, int, str)
    progress = Signal(str)
    finished = Signal(int)

    def __init__(self, search_roots, query, extensions):
        super().__init__()
        self.search_roots = search_roots
        self.query = query.casefold()
        self.extensions = extensions
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True

    def run(self):
        match_count = 0

        for search_root in self.search_roots:
            if self._is_cancelled:
                break

            self.progress.emit(f"Сканирование: {search_root}")

            for root, dirs, files in os.walk(search_root, onerror=self.handle_walk_error):
                if self._is_cancelled:
                    break

                dirs[:] = [dir_name for dir_name in dirs if not dir_name.startswith(".")]

                for file_name in files:
                    if self._is_cancelled:
                        break

                    if not self.matches_extension(file_name):
                        continue

                    full_path = os.path.join(root, file_name)
                    name_matches = self.search_in_file_name(file_name)
                    for snippet in name_matches:
                        self.result_found.emit(full_path, 0, snippet)
                        match_count += 1

                    matches = self.search_in_file(full_path)
                    for line_number, snippet in matches:
                        self.result_found.emit(full_path, line_number, snippet)
                        match_count += 1

        self.finished.emit(match_count)

    def handle_walk_error(self, error):
        return None

    def matches_extension(self, file_name):
        if not self.extensions:
            return True

        lowered_name = file_name.casefold()
        return any(lowered_name.endswith(extension) for extension in self.extensions)

    def search_in_file(self, file_path):
        if file_path.casefold().endswith(".xlsx"):
            return self.search_in_xlsx(file_path)

        matches = []

        try:
            with open(file_path, "r", encoding="utf-8", errors="ignore") as file:
                for line_number, line in enumerate(file, start=1):
                    if self._is_cancelled:
                        break

                    if self.query in line.casefold():
                        snippet = line.strip()
                        if len(snippet) > 140:
                            snippet = f"{snippet[:137]}..."
                        matches.append((line_number, snippet))
                        if len(matches) >= 5:
                            break
        except (OSError, UnicodeError):
            return []

        return matches

    def search_in_file_name(self, file_name):
        if self.query in file_name.casefold():
            return [f"Совпадение в имени файла: {file_name}"]
        return []

    def search_in_xlsx(self, file_path):
        matches = []

        try:
            with zipfile.ZipFile(file_path) as workbook:
                for member_name in workbook.namelist():
                    if self._is_cancelled:
                        break

                    if not member_name.endswith(".xml"):
                        continue

                    if not member_name.startswith(("xl/sharedStrings", "xl/worksheets", "docProps")):
                        continue

                    with workbook.open(member_name) as member:
                        content = member.read().decode("utf-8", errors="ignore")

                    root = ET.fromstring(content)
                    text_chunks = [
                        node.text.strip()
                        for node in root.iter()
                        if node.text and node.text.strip()
                    ]

                    for index, text in enumerate(text_chunks, start=1):
                        if self.query in text.casefold():
                            snippet = text if len(text) <= 140 else f"{text[:137]}..."
                            matches.append((index, snippet))
                            if len(matches) >= 5:
                                return matches
        except (OSError, UnicodeError, zipfile.BadZipFile, ET.ParseError):
            return []

        return matches


class FolderSearchApp(QWidget):
    def __init__(self):
        super().__init__()

        self.search_roots = self.get_search_roots()
        self.search_thread = None
        self.search_worker = None
        self.is_searching = False

        self.setWindowTitle("Network Search")
        self.resize(920, 680)
        self.setup_ui()
        self.apply_styles()
        self.refresh_roots_label()

    def setup_ui(self):
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(24, 24, 24, 24)
        main_layout.setSpacing(18)

        title_label = QLabel("Поиск текста в локальных и сетевых папках")
        title_label.setObjectName("HeroTitle")
        title_font = QFont()
        title_font.setPointSize(18)
        title_font.setBold(True)
        title_label.setFont(title_font)

        subtitle_label = QLabel(
            "Поиск выполняется внутри файлов. Можно искать по локальным дискам "
            "или указать отдельный сетевой путь."
        )
        subtitle_label.setWordWrap(True)
        subtitle_label.setObjectName("SubtitleLabel")

        quick_tip_label = QLabel(
            "Введите текст, при необходимости укажите сетевой путь и сузьте область "
            "через расширения файлов."
        )
        quick_tip_label.setWordWrap(True)
        quick_tip_label.setObjectName("HeroNote")

        main_layout.addWidget(title_label)
        main_layout.addWidget(subtitle_label)
        main_layout.addWidget(quick_tip_label)

        controls_layout = QVBoxLayout()
        controls_layout.setSpacing(14)

        controls_title = QLabel("Параметры поиска")
        controls_title.setObjectName("SectionTitle")

        query_label = QLabel("Текст для поиска")
        query_label.setObjectName("FieldLabel")
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Введите слово или фразу для поиска внутри файлов")
        self.search_input.returnPressed.connect(self.scan_folder)

        network_label = QLabel("Сетевой путь")
        network_label.setObjectName("FieldLabel")
        self.network_path_input = QLineEdit()
        if os.name == "nt":
            self.network_path_input.setPlaceholderText(r"Например: \\server\shared")
        else:
            self.network_path_input.setPlaceholderText("Например: /Volumes/Shared")

        extensions_label = QLabel("Расширения файлов")
        extensions_label.setObjectName("FieldLabel")
        self.extensions_input = QLineEdit()
        self.extensions_input.setPlaceholderText("Например: .txt,.py,.log или оставить пустым")

        self.search_hint_label = QLabel(
            "Если сетевой путь не указан, поиск выполняется только по доступным папкам этого компьютера."
        )
        self.search_hint_label.setWordWrap(True)
        self.search_hint_label.setObjectName("HintLabel")

        self.path_label = QLabel()
        self.path_label.setWordWrap(True)
        self.path_label.setObjectName("PathLabel")

        button_row = QHBoxLayout()
        button_row.setSpacing(10)

        self.scan_button = QPushButton("Начать поиск")
        self.scan_button.clicked.connect(self.scan_folder)

        self.stop_button = QPushButton("Остановить")
        self.stop_button.clicked.connect(self.stop_search)
        self.stop_button.setEnabled(False)

        self.open_button = QPushButton("Открыть выбранное")
        self.open_button.clicked.connect(self.open_selected_item)

        button_row.addWidget(self.scan_button)
        button_row.addWidget(self.stop_button)
        button_row.addWidget(self.open_button)

        controls_layout.addWidget(controls_title)
        controls_layout.addWidget(query_label)
        controls_layout.addWidget(self.search_input)
        controls_layout.addWidget(network_label)
        controls_layout.addWidget(self.network_path_input)
        controls_layout.addWidget(extensions_label)
        controls_layout.addWidget(self.extensions_input)
        controls_layout.addWidget(self.search_hint_label)
        controls_layout.addWidget(self.path_label)
        controls_layout.addLayout(button_row)
        main_layout.addLayout(controls_layout)

        results_layout = QVBoxLayout()
        results_layout.setSpacing(12)

        results_title = QLabel("Результаты")
        results_title.setObjectName("SectionTitle")
        self.status_label = QLabel("Введите текст и начните поиск.")
        self.status_label.setObjectName("StatusLabel")

        self.results_list = QListWidget()
        self.results_list.itemDoubleClicked.connect(self.open_selected_item)

        results_layout.addWidget(results_title)
        results_layout.addWidget(self.status_label)
        results_layout.addWidget(self.results_list)
        main_layout.addLayout(results_layout)

        self.setLayout(main_layout)

    def apply_styles(self):
        self.setStyleSheet(
            """
            QWidget {
                background: #eef3f8;
                color: #18253a;
                font-size: 14px;
            }
            QLabel {
                background: transparent;
            }
            QLabel#HeroTitle {
                color: #16253d;
            }
            QLabel#SubtitleLabel, QLabel#HeroNote {
                color: #56637a;
            }
            QLabel#PathLabel, QLabel#StatusLabel, QLabel#HintLabel {
                color: #56637a;
            }
            QLabel#SectionTitle {
                color: #16253d;
                font-size: 16px;
                font-weight: 700;
            }
            QLabel#FieldLabel {
                color: #20314f;
                font-weight: 600;
            }
            QLineEdit {
                background: #fbfdff;
                border: 1px solid #c9d7e8;
                border-radius: 14px;
                padding: 13px 14px;
                selection-background-color: #2d6cdf;
            }
            QLineEdit:focus {
                border: 1px solid #2d6cdf;
                background: #ffffff;
            }
            QPushButton {
                background: #2d6cdf;
                color: white;
                border: none;
                border-radius: 14px;
                padding: 13px 18px;
                font-weight: 600;
            }
            QPushButton:hover {
                background: #2159bb;
            }
            QPushButton:pressed {
                background: #1b4a9c;
            }
            QPushButton:disabled {
                background: #b9c6db;
                color: #eef3fb;
            }
            QListWidget {
                background: #fbfcff;
                border: 1px solid #d9e2f0;
                border-radius: 16px;
                padding: 10px;
                outline: none;
            }
            QListWidget {
                font-size: 13px;
            }
            QListWidget::item {
                padding: 12px;
                margin-bottom: 6px;
                border: 1px solid #e4ebf4;
                border-radius: 12px;
                background: #ffffff;
            }
            QListWidget::item:selected {
                background: #dce9ff;
                color: #10213f;
                border: 1px solid #b7cff7;
            }
            """
        )

    def refresh_roots_label(self):
        roots_text = ", ".join(self.search_roots) if self.search_roots else "не найдены"
        self.path_label.setText(f"Автоматические области поиска: {roots_text}")

    def get_search_roots(self):
        if sys.platform == "darwin":
            volumes_path = "/Volumes"
            if os.path.isdir(volumes_path):
                mounted_volumes = [
                    os.path.join(volumes_path, item)
                    for item in os.listdir(volumes_path)
                    if os.path.isdir(os.path.join(volumes_path, item))
                ]
                if mounted_volumes:
                    return mounted_volumes
            return ["/"]

        if os.name == "nt":
            available_drives = []
            for letter in string.ascii_uppercase:
                drive = f"{letter}:\\"
                if os.path.exists(drive):
                    available_drives.append(drive)
            return available_drives or [r"C:\\"]

        return ["/"]

    def build_search_roots(self):
        search_roots = list(self.search_roots)
        custom_path = self.network_path_input.text().strip()
        if custom_path:
            if os.path.exists(custom_path):
                if custom_path not in search_roots:
                    search_roots.insert(0, custom_path)
            else:
                QMessageBox.warning(
                    self,
                    "Путь недоступен",
                    "Указанный сетевой путь не найден или недоступен.",
                )
                return []

        return search_roots

    def parse_extensions(self):
        raw_extensions = self.extensions_input.text().strip()
        if not raw_extensions:
            return []

        extensions = []
        for item in raw_extensions.split(","):
            extension = item.strip().casefold()
            if not extension:
                continue
            if not extension.startswith("."):
                extension = f".{extension}"
            extensions.append(extension)
        return extensions

    def scan_folder(self):
        if self.is_searching:
            return

        query = self.search_input.text().strip()
        if not query:
            QMessageBox.warning(self, "Ошибка", "Введите слово или фразу для поиска.")
            return

        search_roots = self.build_search_roots()
        if not search_roots:
            return

        extensions = self.parse_extensions()

        self.results_list.clear()
        self.status_label.setText("Подготавливаю поиск...")
        self.set_search_state(True)

        self.search_thread = QThread()
        self.search_worker = SearchWorker(search_roots, query, extensions)
        self.search_worker.moveToThread(self.search_thread)

        self.search_thread.started.connect(self.search_worker.run)
        self.search_worker.result_found.connect(self.add_result_item)
        self.search_worker.progress.connect(self.update_progress)
        self.search_worker.finished.connect(self.finish_search)
        self.search_worker.finished.connect(self.search_thread.quit)
        self.search_worker.finished.connect(self.search_worker.deleteLater)
        self.search_thread.finished.connect(self.search_thread.deleteLater)
        self.search_thread.finished.connect(self.cleanup_search_thread)

        self.search_thread.start()

    def set_search_state(self, is_searching):
        self.is_searching = is_searching
        self.scan_button.setEnabled(not is_searching)
        self.stop_button.setEnabled(is_searching)
        self.search_input.setEnabled(not is_searching)
        self.network_path_input.setEnabled(not is_searching)
        self.extensions_input.setEnabled(not is_searching)

    def update_progress(self, message):
        self.status_label.setText(message)

    def stop_search(self):
        if self.search_worker:
            self.search_worker.cancel()
            self.status_label.setText("Останавливаю поиск...")

    def finish_search(self, match_count):
        if self.search_worker and self.search_worker._is_cancelled:
            self.status_label.setText(f"Поиск остановлен. Найдено вхождений: {match_count}")
        elif match_count:
            self.status_label.setText(f"Найдено вхождений: {match_count}")
        else:
            self.status_label.setText("Вхождения не найдены.")

        self.set_search_state(False)

        QMessageBox.information(
            self,
            "Поиск завершен",
            f"Найдено вхождений: {match_count}",
        )

    def cleanup_search_thread(self):
        self.search_thread = None
        self.search_worker = None

    def add_result_item(self, full_path, line_number, snippet):
        item = QListWidgetItem(f"[Строка {line_number}] {full_path}\n{snippet}")
        item.setData(Qt.UserRole, full_path)
        self.results_list.addItem(item)

    def open_selected_item(self, item=None):
        selected_item = item or self.results_list.currentItem()
        if not selected_item:
            QMessageBox.information(self, "Нет выбора", "Сначала выберите файл или папку.")
            return

        path = selected_item.data(Qt.UserRole)
        if not path or not os.path.exists(path):
            QMessageBox.warning(self, "Ошибка", "Выбранный путь больше недоступен.")
            return

        try:
            if os.name == "nt":
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.run(["open", path], check=False)
            else:
                subprocess.run(["xdg-open", path], check=False)
        except Exception as error:
            QMessageBox.warning(self, "Ошибка открытия", str(error))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FolderSearchApp()
    window.show()
    sys.exit(app.exec())
