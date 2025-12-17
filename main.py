import os
import time
import traceback
from dotenv import load_dotenv
from modules.mail_reader import MailReader
from modules.parser import EmailParser
from modules.doc_generator import DocGenerator
from modules.sender import EmailSender
from modules.logger import OperationLogger

# Загрузка переменных окружения
load_dotenv()

# Конфигурация из .env
MAIL_LOGIN = os.getenv("MAIL_LOGIN")
MAIL_PASSWORD = os.getenv("MAIL_PASSWORD")
SECURITY_EMAIL = os.getenv("SECURITY_EMAIL")

IMAP_SERVER = os.getenv("IMAP_SERVER", "imap.mail.ru")
IMAP_PORT = int(os.getenv("IMAP_PORT", "993"))
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.mail.ru")
SMTP_PORT = int(os.getenv("SMTP_PORT", "465"))

TEMPLATE_PATH = os.getenv("TEMPLATE_PATH", "templates/Шаблон_Служебной_записки.docx")
OUTPUT_DIR = os.getenv("OUTPUT_DIR", "output")


def main():
    # Инициализация компонентов
    mail_reader = MailReader(login=MAIL_LOGIN, password=MAIL_PASSWORD, server=IMAP_SERVER, port=IMAP_PORT)
    parser = EmailParser()
    doc_generator = DocGenerator(template_path=TEMPLATE_PATH, output_dir=OUTPUT_DIR)
    sender = EmailSender(login=MAIL_LOGIN, password=MAIL_PASSWORD, smtp_server=SMTP_SERVER, smtp_port=SMTP_PORT)
    operation_logger = OperationLogger()

    try:
        # Подключение к IMAP
        if not mail_reader.connect():
            operation_logger.log_operation(
                status="Ошибка",
                error_message="Не удалось подключиться к IMAP-серверу"
            )
            return

        # Получение непрочитанных писем
        emails = mail_reader.fetch_unread_emails()
        if not emails:
            print("Новых писем нет.")
            return

        for email in emails:
            print(f"\n➡ Обработка письма от: {email['from']} | Тема: {email['subject']}")
            license_plate = None
            output_file = None

            try:
                # Парсинг данных
                parsed_data = parser.parse(email["body"], email["subject"])

                # Обязательные поля
                license_plate = parsed_data.get("license_plate")
                visit_date = parsed_data.get("visit_date")

                if not license_plate or not visit_date:
                    raise ValueError("Отсутствуют обязательные поля: гос. номер или дата")

                # Генерация документа
                output_file = doc_generator.generate(parsed_data)
                if not output_file:
                    raise RuntimeError("Не удалось сгенерировать служебную записку")

                # Отправка в СБ
                success = sender.send_pass_request(
                    to_email=SECURITY_EMAIL,
                    license_plate=license_plate,
                    visit_date=visit_date,
                    docx_path=output_file
                )
                if not success:
                    raise RuntimeError("Не удалось отправить письмо в СБ")

                # Успех: помечаем письмо как прочитанное и логируем
                mail_reader.mark_as_read(email["uid"])
                operation_logger.log_operation(
                    status="Успех",
                    license_plate=license_plate,
                    output_filename=os.path.basename(output_file)
                )
                print(f"✅ Успешно обработано: {license_plate}")

            except Exception as e:
                error_msg = f"{type(e).__name__}: {str(e)}"
                print(f"❌ Ошибка: {error_msg}")
                operation_logger.log_operation(
                    status="Ошибка",
                    license_plate=license_plate,
                    output_filename=os.path.basename(output_file) if output_file else None,
                    error_message=error_msg[:500]
                )
                # В случае ошибки — НЕ помечаем письмо как прочитанное,
                # чтобы можно было повторно обработать после исправления

    except Exception as e:
        error_full = traceback.format_exc()
        print(f"КРИТИЧЕСКАЯ ОШИБКА: {error_full}")
        operation_logger.log_operation(
            status="Ошибка",
            error_message=f"Критическая ошибка в основном цикле: {str(e)}"[:500]
        )
    finally:
        # Закрытие ресурсов
        if 'mail_reader' in locals():
            mail_reader.disconnect()
        if 'operation_logger' in locals():
            operation_logger.close()


if __name__ == "__main__":
    main()