import os
import imaplib
import email
from email.header import decode_header
from typing import List, Optional, Tuple, Dict
import time
import logging
from datetime import datetime


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("logs/mail_reader.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("MailReader")


class MailReader:
    def __init__(self, login: str, password: str, server: str = "imap.mail.ru", port: int = 993):
        self.login = login
        self.password = password
        self.server = server
        self.port = port
        self.mail = None

    def connect(self, max_retries: int = 3, retry_delay: int = 10) -> bool:
        """Устанавливает IMAP-соединение с повторными попытками."""
        for attempt in range(1, max_retries + 1):
            try:
                self.mail = imaplib.IMAP4_SSL(self.server, self.port)
                self.mail.login(self.login, self.password)
                logger.info("Успешное подключение к IMAP-серверу.")
                return True
            except Exception as e:
                logger.error(f"Попытка {attempt}/{max_retries} не удалась: {e}")
                if attempt < max_retries:
                    time.sleep(retry_delay)
                else:
                    logger.critical("Не удалось подключиться к почтовому серверу после всех попыток.")
                    return False

    def disconnect(self):
        """Закрывает соединение с почтовым сервером."""
        if self.mail:
            try:
                self.mail.close()
                self.mail.logout()
                logger.info("Соединение с IMAP-сервером закрыто.")
            except Exception as e:
                logger.warning(f"Ошибка при закрытии соединения: {e}")

    def _decode_mime_word(self, encoded: str) -> str:
        """Декодирует заголовки email (например, темы)."""
        decoded_fragments = decode_header(encoded)
        parts = []
        for fragment, encoding in decoded_fragments:
            if isinstance(fragment, bytes):
                parts.append(fragment.decode(encoding or "utf-8", errors="replace"))
            else:
                parts.append(fragment)
        return "".join(parts)

    def _get_email_body(self, msg) -> str:
        """Извлекает текстовое тело письма (предпочтительно plain, fallback — html)."""
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                if "attachment" in content_disposition:
                    continue
                if content_type == "text/plain":
                    payload = part.get_payload(decode=True)
                    if payload:
                        body = payload.decode(part.get_content_charset() or "utf-8", errors="replace")
                        break
                elif content_type == "text/html" and not body:
                    payload = part.get_payload(decode=True)
                    if payload:
                        html_body = payload.decode(part.get_content_charset() or "utf-8", errors="replace")
                        body = html_body
        else:
            payload = msg.get_payload(decode=True)
            if payload:
                body = payload.decode(msg.get_content_charset() or "utf-8", errors="replace")
        return body.strip()

    def fetch_unread_emails(self) -> List[Dict]:
        """
        Получает все непрочитанные письма из INBOX.
        Возвращает список словарей с полями:
        - uid (str)
        - from (str)
        - subject (str)
        - date (datetime)
        - body (str)
        """
        if not self.mail:
            logger.error("Нет активного соединения с IMAP.")
            return []

        try:
            self.mail.select("INBOX")
            status, messages = self.mail.search(None, "UNSEEN")
            if status != "OK":
                logger.error("Не удалось выполнить поиск непрочитанных писем.")
                return []

            email_ids = messages[0].split()
            if not email_ids:
                logger.info("Нет новых непрочитанных писем.")
                return []

            emails = []
            for email_id in email_ids:
                try:
                    # Получаем письмо по UID
                    status, msg_data = self.mail.fetch(email_id, "(RFC822)")
                    if status != "OK":
                        logger.warning(f"Не удалось получить письмо с ID {email_id}.")
                        continue

                    raw_email = msg_data[0][1]
                    msg = email.message_from_bytes(raw_email)

                    # Отправитель
                    from_ = msg.get("From", "")
                    # Тема
                    subject_encoded = msg.get("Subject", "")
                    subject = self._decode_mime_word(subject_encoded)
                    # Дата
                    date_str = msg.get("Date", "")
                    try:
                        # email.utils.parsedate_to_datetime безопаснее
                        email_date = email.utils.parsedate_to_datetime(date_str)
                    except Exception:
                        email_date = datetime.now()
                        logger.warning(f"Некорректная дата в письме: {date_str}. Использована текущая дата.")

                    # Тело
                    body = self._get_email_body(msg)

                    emails.append({
                        "uid": email_id.decode(),
                        "from": from_,
                        "subject": subject,
                        "date": email_date,
                        "body": body
                    })

                except Exception as e:
                    logger.error(f"Ошибка при обработке письма {email_id}: {e}")
                    continue

            logger.info(f"Найдено {len(emails)} новых непрочитанных писем.")
            return emails

        except Exception as e:
            logger.error(f"Ошибка при получении писем: {e}")
            return []

    def mark_as_read(self, email_uid: str) -> bool:
        """Помечает письмо как прочитанное по его UID."""
        try:
            self.mail.store(email_uid, '+FLAGS', '\\Seen')
            logger.debug(f"Письмо UID={email_uid} помечено как прочитанное.")
            return True
        except Exception as e:
            logger.error(f"Не удалось пометить письмо UID={email_uid} как прочитанное: {e}")
            return False