from pathlib import Path
from typing import Any, List, Optional

from PyQt5.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QPushButton,
    QFileDialog,
    QMessageBox,
    QComboBox,
    QLineEdit,
    QLabel,
)

from macros.excel.work_exel import ExcelService, FilterRequest


class App(QWidget):
    """UI для фильтрации Excel.

    Сценарий:
        1) Открыть входной файл.
        2) Выбрать столбец и ввести значение.
        3) Нажать "Фильтровать".
        4) Нажать "Сохранить" (запись результата в .xlsx).
    """

    def __init__(self, excel_service: ExcelService, logger: Optional[object] = None) -> None:
        super().__init__()

        self.excel_service = excel_service
        self.logger = logger

        self.setWindowTitle("Excel фильтр (минимум)")
        self.resize(520, 240)

        self.input_file_path: Optional[Path] = None

        self.filtered_headers: Optional[List[str]] = None
        self.filtered_rows: Optional[List[List[Any]]] = None

        self.open_file_button = QPushButton("1) Открыть Excel")
        self.filter_button = QPushButton("2) Фильтровать")
        self.save_button = QPushButton("3) Сохранить")

        self.filter_button.setEnabled(False)
        self.save_button.setEnabled(False)

        self.input_file_label = QLabel("Файл: (не выбран)")
        self.filter_column_combo = QComboBox()
        self.filter_value_input = QLineEdit()

        self._build_layout()
        self._bind_signals()

    def _build_layout(self) -> None:
        """Собирает layout окна."""
        layout = QVBoxLayout()

        layout.addWidget(self.open_file_button)
        layout.addWidget(self.input_file_label)

        layout.addWidget(QLabel("Столбец для фильтра:"))
        layout.addWidget(self.filter_column_combo)

        layout.addWidget(QLabel("Значение для фильтра:"))
        layout.addWidget(self.filter_value_input)

        layout.addWidget(self.filter_button)
        layout.addWidget(self.save_button)

        self.setLayout(layout)

    def _bind_signals(self) -> None:
        """Подключает обработчики кнопок."""
        self.open_file_button.clicked.connect(self.open_input_file)
        self.filter_button.clicked.connect(self.apply_filter)
        self.save_button.clicked.connect(self.save_filtered_result)

    def open_input_file(self) -> None:
        """Открывает диалог выбора файла и загружает заголовки таблицы.
        """
        selected_path, _ = QFileDialog.getOpenFileName(
            self,
            "Открыть Excel",
            "",
            "Excel files (*.xlsx *.xlsm *.xls)",
        )
        if not selected_path:
            return

        self.input_file_path = Path(selected_path)
        self.input_file_label.setText(f"Файл: {self.input_file_path.name}")

        self.filtered_headers = None
        self.filtered_rows = None
        self.save_button.setEnabled(False)

        try:
            headers = self.excel_service.read_headers(self.input_file_path)

            self.filter_column_combo.clear()
            self.filter_column_combo.addItems(headers)

            self.filter_button.setEnabled(True)

            if self.logger:
                self.logger.info("Открыт файл: %s", self.input_file_path.name)

        except Exception as exc:
            if self.logger:
                self.logger.exception("Не удалось прочитать заголовки")

            QMessageBox.critical(
                self,
                "Ошибка",
                f"Не удалось прочитать заголовки:\n{exc}",
            )
            self.filter_button.setEnabled(False)

    def apply_filter(self) -> None:
        """Фильтрует строки и сохраняет результат в памяти приложения.

        Шаг НЕ сохраняет файл на диск. Он только:
            - формирует FilterRequest (без output_path),
            - вызывает ExcelService.filter_rows(),
            - сохраняет headers/rows в self.filtered_headers/self.filtered_rows,
            - включает кнопку "Сохранить".
        """
        if self.input_file_path is None:
            QMessageBox.warning(self, "Проверка", "Сначала выберите Excel-файл.")
            return

        filter_column = self.filter_column_combo.currentText().strip()
        filter_value = self.filter_value_input.text()

        if not filter_column:
            QMessageBox.warning(self, "Проверка", "Выберите столбец для фильтра.")
            return

        request = FilterRequest(
            input_path=self.input_file_path,
            filter_column=filter_column,
            filter_value=filter_value,
            output_path=None,
        )

        try:
            headers, rows = self.excel_service.filter_rows(request)

            self.filtered_headers = headers
            self.filtered_rows = rows
            self.save_button.setEnabled(True)

            QMessageBox.information(self, "Фильтрация", f"Найдено строк: {len(rows)}")

            if self.logger:
                self.logger.info(
                    "Фильтрация выполнена: column=%s value=%s rows=%s",
                    filter_column,
                    filter_value,
                    len(rows),
                )

        except Exception as exc:
            if self.logger:
                self.logger.exception("Ошибка фильтрации")

            QMessageBox.critical(self, "Ошибка", str(exc))
            self.save_button.setEnabled(False)

    def save_filtered_result(self) -> None:
        """Сохраняет уже отфильтрованный результат в .xlsx.
        """
        if not self.filtered_headers or self.filtered_rows is None:
            QMessageBox.warning(self, "Проверка", "Сначала нажмите «Фильтровать».")
            return

        selected_path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить результат",
            "filtered_result.xlsx",
            "Excel files (*.xlsx)",
        )
        if not selected_path:
            return

        output_path = Path(selected_path)
        if output_path.suffix.lower() != ".xlsx":
            output_path = output_path.with_suffix(".xlsx")

        if output_path.exists():
            answer = QMessageBox.question(
                self,
                "Перезаписать?",
                f"Файл уже существует:\n{output_path.name}\n\nПерезаписать?",
                QMessageBox.Yes | QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                return

        try:
            self.excel_service.save_xlsx(output_path, self.filtered_headers, self.filtered_rows)

            QMessageBox.information(
                self,
                "Готово",
                f"Сохранено строк: {len(self.filtered_rows)}\nФайл: {output_path.name}",
            )

            if self.logger:
                self.logger.info(
                    "Результат сохранён: %s (rows=%s)",
                    output_path.name,
                    len(self.filtered_rows),
                )

        except Exception as exc:
            if self.logger:
                self.logger.exception("Ошибка сохранения")

            QMessageBox.critical(self, "Ошибка", str(exc))