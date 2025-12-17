import re
import logging
from typing import Dict, Optional

# Настройка логирования (можно позже вынести в общий logger)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("logs/parser.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("Parser")


class EmailParser:
    def __init__(self):
        self.patterns = {
            "vehicle": re.compile(
                r"(?:Автомобиль|Марка(?:\s+ТС)?|Транспортное средство)[:\s]+([А-ЯЁа-яёa-zA-Z0-9\s\-]+?)(?:\n|—|$)", 
                re.IGNORECASE | re.DOTALL
            ),
            "license_plate": re.compile(
                r"[А-ЯA-Z]\d{3}[А-ЯA-Z]{2}\d{2,3}", re.IGNORECASE
            ),
            "driver_name": re.compile(
                r"(?:Водитель|ФИО водителя|Водитель:\s*)([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)",
                re.IGNORECASE
            ),
            "driver_phone": re.compile(
                r"(\+7\s*\d{3}[-\s]?\d{3}[-\s]?\d{2}[-\s]?\d{2})|"
                r"(\+7\s*\(\d{3}\)\s*\d{3}-\d{2}-\d{2})",
                re.IGNORECASE
            ),
            "visit_date": re.compile(
                r"(\d{2}\.\d{2}\.\d{4})|(\d{2}\.\d{2}\.20\d{2})", re.IGNORECASE
            ),
            "time_range": re.compile(
                r"с\s*(\d{1,2}:\d{2})\s*до\s*(\d{1,2}:\d{2})", re.IGNORECASE
            ),
            "visit_purpose": re.compile(
                r"(?:Цель визита|Цель:|Цель заезда|Причина:)\s*(.+?)(?:\n|$)", re.IGNORECASE
            )
        }

    def _extract_vehicle(self, text: str) -> Optional[str]:
        match = self.patterns["vehicle"].search(text)
        if match:
            return match.group(1).strip()
        return None

    def _extract_license_plate(self, text: str) -> Optional[str]:
        match = self.patterns["license_plate"].search(text)
        if match:
            return match.group(0).upper()
        return None

    def _extract_driver_name(self, text: str) -> Optional[str]:
        match = self.patterns["driver_name"].search(text)
        if match:
            return match.group(1).strip()
        # Альтернатива: если ФИО идёт после номера или рядом с телефоном
        # (можно расширить при необходимости)
        return None

    def _extract_driver_phone(self, text: str) -> Optional[str]:
        match = self.patterns["driver_phone"].search(text)
        if match:
            return match.group(0).replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
        return None

    def _extract_visit_date(self, text: str) -> Optional[str]:
        match = self.patterns["visit_date"].search(text)
        if match:
            return match.group(0)
        return None

    def _extract_time_range(self, text: str) -> Dict[str, Optional[str]]:
        match = self.patterns["time_range"].search(text)
        if match:
            return {
                "start": match.group(1),
                "end": match.group(2)
            }
        return {"start": None, "end": None}

    def _extract_visit_purpose(self, text: str) -> Optional[str]:
        match = self.patterns["visit_purpose"].search(text)
        if match:
            return match.group(1).strip().rstrip(".")
        # Если нет явной метки — ищем после временного интервала или даты
        lines = text.split("\n")
        purpose = None
        for i, line in enumerate(lines):
            if re.search(r"\d{2}:\d{2}", line):
                if i + 1 < len(lines):
                    next_line = lines[i + 1].strip()
                    if next_line and not re.search(r"[А-Яа-яЁё]", next_line):
                        continue
                    if len(next_line) > 10:
                        purpose = next_line
                        break
        if not purpose:
            # Или последняя непустая строка, если она похожа на цель
            for line in reversed(lines):
                line = line.strip()
                if line and len(line) > 15 and not re.search(r"[\d@+]", line):
                    purpose = line
                    break
        return purpose.rstrip(".") if purpose else None

    def parse(self, email_body: str, subject: str = "") -> Dict[str, Optional[str]]:
        """
        Парсит тело письма и возвращает словарь с извлечёнными данными.
        Все поля опциональны. При отсутствии — None.
        """
        full_text = f"{subject}\n{email_body}".strip()

        vehicle = self._extract_vehicle(full_text)
        license_plate = self._extract_license_plate(full_text)
        driver_name = self._extract_driver_name(full_text)
        driver_phone = self._extract_driver_phone(full_text)
        visit_date = self._extract_visit_date(full_text)
        time_range = self._extract_time_range(full_text)
        visit_purpose = self._extract_visit_purpose(full_text)

        result = {
            "vehicle": vehicle,
            "license_plate": license_plate,
            "driver_name": driver_name,
            "driver_phone": driver_phone,
            "visit_date": visit_date,
            "visit_time_start": time_range["start"],
            "visit_time_end": time_range["end"],
            "visit_purpose": visit_purpose
        }

        # Логируем отсутствие критических полей
        if not license_plate:
            logger.warning("Не удалось извлечь гос. номер.")
        if not visit_date:
            logger.warning("Не удалось извлечь дату доставки.")
        if not result["visit_time_start"] or not result["visit_time_end"]:
            logger.warning("Не удалось извлечь полный временной интервал.")

        return result