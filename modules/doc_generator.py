import os
import logging
from datetime import datetime
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from typing import Dict, Optional


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("logs/doc_generator.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("DocGenerator")


class DocGenerator:
    def __init__(self, template_path: str, output_dir: str):
        self.template_path = template_path
        self.output_dir = output_dir
        os.makedirs(self.output_dir, exist_ok=True)

    def _format_date(self, date_str: Optional[str]) -> str:
        """Преобразует дату в формат ДД.ММ.ГГГГ.
        Поддерживает строки вида '05.12.2025' → '05.12.2025'."""
        if not date_str:
            return datetime.now().strftime("%d.%m.%Y")
        # Убедимся, что формат корректный
        try:
            date_obj = datetime.strptime(date_str, "%d.%m.%Y")
            return date_obj.strftime("%d.%m.%Y")
        except ValueError:
            logger.warning(f"Некорректный формат даты: {date_str}. Используется текущая дата.")
            return datetime.now().strftime("%d.%m.%Y")

    def _safe_get(self, data: Dict, key: str, default: str = "") -> str:
        """Безопасное извлечение значения из словаря."""
        value = data.get(key)
        return str(value).strip() if value else default

    def generate(self, parsed_data: Dict[str, Optional[str]]) -> Optional[str]:
        """
        Генерирует служебную записку и возвращает путь к созданному файлу.
        Если возникла ошибка — возвращает None.
        """
        try:
            # Загрузка шаблона
            if not os.path.exists(self.template_path):
                logger.error(f"Шаблон не найден: {self.template_path}")
                return None

            doc = Document(self.template_path)

            # Текущая дата (дата формирования записки)
            current_date = datetime.now().strftime("%d.%m.%Y")

            # Дата визита
            visit_date = self._format_date(parsed_data.get("visit_date"))

            # Подготовка данных для замены
            replacements = {
                "{current_date}": current_date,
                "{vehicle}": self._safe_get(parsed_data, "vehicle", "Автомобиль не указан"),
                "{license_plate}": self._safe_get(parsed_data, "license_plate", "Номер не указан"),
                "{driver_name}": self._safe_get(parsed_data, "driver_name", "ФИО не указано"),
                "{driver_phone}": self._safe_get(parsed_data, "driver_phone", "Телефон не указан"),
                "{visit_date}": visit_date,
                "{visit_time_start}": self._safe_get(parsed_data, "visit_time_start", "время не указано"),
                "{visit_time_end}": self._safe_get(parsed_data, "visit_time_end", "время не указано"),
                "{visit_purpose}": self._safe_get(parsed_data, "visit_purpose", "Цель не указана")
            }

            # Замена плейсхолдеров во всём документе
            for paragraph in doc.paragraphs:
                for placeholder, value in replacements.items():
                    if placeholder in paragraph.text:
                        paragraph.text = paragraph.text.replace(placeholder, value)
                        # Сохраняем выравнивание (если важно)
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

            # Формирование имени файла
            license_plate = replacements["{license_plate}"].replace(" ", "")
            visit_date_short = visit_date.replace(".", "")[2:]  # -> 251205
            filename = f"СЗ_{license_plate}_{visit_date_short}.docx"
            output_path = os.path.join(self.output_dir, filename)

            # Сохранение
            doc.save(output_path)
            logger.info(f"Служебная записка сохранена: {output_path}")
            return output_path

        except Exception as e:
            logger.error(f"Ошибка при генерации документа: {e}")
            return None