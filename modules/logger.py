import os
import csv
from datetime import datetime
from typing import Optional

# Путь к папке логов (можно вынести в конфиг, но для простоты — жёстко задан)
LOG_DIR = "logs"
os.makedirs(LOG_DIR, exist_ok=True)


class OperationLogger:
    def __init__(self):
        self._current_date = None
        self._file_path = None
        self._file = None
        self._writer = None

    def _get_log_filename(self) -> str:
        """Возвращает путь к CSV-файлу текущего дня."""
        today = datetime.now().strftime("%Y-%m-%d")
        return os.path.join(LOG_DIR, f"log_{today}.csv")

    def _open_writer(self):
        """Открывает CSV-файл для записи (создаёт с заголовком, если нужно)."""
        log_path = self._get_log_filename()
        file_exists = os.path.isfile(log_path)

        self._file = open(log_path, mode="a", newline="", encoding="utf-8")
        self._writer = csv.writer(self._file, delimiter=";")  # точка с запятой — для совместимости с Excel

        # Записываем заголовок, если файл новый
        if not file_exists:
            self._writer.writerow([
                "timestamp",
                "status",
                "license_plate",
                "output_filename",
                "error_message"
            ])
            self._file.flush()

    def log_operation(
        self,
        status: str,
        license_plate: Optional[str] = None,
        output_filename: Optional[str] = None,
        error_message: Optional[str] = None
    ):
        """
        Записывает событие в CSV-журнал.
        :param status: "Успех" или "Ошибка"
        :param license_plate: Гос. номер ТС
        :param output_filename: Имя сгенерированного .docx
        :param error_message: Причина ошибки (только при статусе "Ошибка")
        """
        timestamp = datetime.now().isoformat(sep=" ", timespec="seconds")

        # Обновляем файл, если сменился день
        today_file = self._get_log_filename()
        if self._file_path != today_file:
            if self._file:
                self._file.close()
            self._file_path = today_file
            self._open_writer()

        self._writer.writerow([
            timestamp,
            status,
            license_plate or "",
            output_filename or "",
            (error_message or "")[:500]  # ограничение длины на случай очень длинных ошибок
        ])
        self._file.flush()

    def close(self):
        """Закрывает файл (рекомендуется вызывать при завершении работы)."""
        if self._file:
            self._file.close()