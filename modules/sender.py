import os
import smtplib
import logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from typing import Optional


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("logs/sender.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("Sender")


class EmailSender:
    def __init__(self, login: str, password: str, smtp_server: str = "smtp.mail.ru", smtp_port: int = 465):
        self.login = login
        self.password = password
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port

    def send_pass_request(
        self,
        to_email: str,
        license_plate: str,
        visit_date: str,
        docx_path: str
    ) -> bool:
        """
        Отправляет служебную записку в СБ.
        Возвращает True при успешной отправке, иначе False.
        """
        try:
            # Тема письма
            subject = f"Пропуск для {license_plate} на {visit_date}"

            # Тело письма
            body = "Уважаемый Иван Сергеевич, прошу подготовить пропуск. Служебная записка во вложении."

            # Создание сообщения
            msg = MIMEMultipart()
            msg["From"] = self.login
            msg["To"] = to_email
            msg["Subject"] = subject
            msg.attach(MIMEText(body, "plain", "utf-8"))

            # Прикрепление файла
            if not os.path.exists(docx_path):
                logger.error(f"Файл для вложения не найден: {docx_path}")
                return False

            with open(docx_path, "rb") as attachment:
                part = MIMEBase("application", "vnd.openxmlformats-officedocument.wordprocessingml.document")
                part.set_payload(attachment.read())
            encoders.encode_base64(part)
            filename = os.path.basename(docx_path)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename={filename}"
            )
            msg.attach(part)

            # Отправка через SMTP SSL
            with smtplib.SMTP_SSL(self.smtp_server, self.smtp_port) as server:
                server.login(self.login, self.password)
                server.sendmail(self.login, to_email, msg.as_string())

            logger.info(f"Письмо успешно отправлено на {to_email}. Вложение: {filename}")
            return True

        except smtplib.SMTPAuthenticationError:
            logger.error("Ошибка аутентификации SMTP: проверьте логин и пароль (используйте пароль для внешних приложений от Mail.ru).")
            return False
        except smtplib.SMTPRecipientsRefused:
            logger.error(f"Почтовый сервер отказался принимать письмо для получателя {to_email}.")
            return False
        except smtplib.SMTPException as e:
            logger.error(f"Ошибка SMTP: {e}")
            return False
        except Exception as e:
            logger.error(f"Неожиданная ошибка при отправке письма: {e}")
            return False