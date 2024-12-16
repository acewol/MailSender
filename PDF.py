import os
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

def send_pdf_emails(folder_path, excel_path, smtp_server, smtp_port, sender_email, sender_password):
    """
    Рассылает PDF файлы пользователям на основе Excel таблицы.

    :param folder_path: Путь к папке с PDF файлами
    :param excel_path: Путь к Excel файлу с колонками 'ФИО' и 'Email'
    :param smtp_server: Адрес SMTP сервера
    :param smtp_port: Порт SMTP сервера
    :param sender_email: Email отправителя
    :param sender_password: Пароль от email отправителя
    """
    try:
        # Загружаем Excel файл
        data = pd.read_excel(excel_path)

        # Проверяем наличие необходимых колонок
        if not {'ФИО', 'Email'}.issubset(data.columns):
            raise ValueError("Excel файл должен содержать колонки 'ФИО' и 'Email'")

        # Устанавливаем соединение с SMTP сервером
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()  # Шифрование соединения
            server.login(sender_email, sender_password)

            # Обходим всех пользователей из Excel файла
            for _, row in data.iterrows():
                full_name = row['ФИО']
                recipient_email = row['Email']

                # Определяем путь к PDF файлу
                pdf_file_name = f"{full_name}.pdf"
                pdf_file_path = os.path.join(folder_path, pdf_file_name)

                # Проверяем, существует ли файл
                if not os.path.exists(pdf_file_path):
                    print(f"Файл {pdf_file_name} не найден. Пропускаем...")
                    continue

                # Формируем письмо
                msg = MIMEMultipart()
                msg['From'] = sender_email
                msg['To'] = recipient_email
                msg['Subject'] = "Ваш документ"

                body = f"Уважаемый(ая) {full_name},\n\nВо вложении вы найдете ваш документ."
                msg.attach(MIMEText(body, 'plain'))

                # Прикрепляем PDF файл
                with open(pdf_file_path, 'rb') as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header(
                        'Content-Disposition',
                        f'attachment; filename={pdf_file_name}'
                    )
                    msg.attach(part)

                # Отправляем письмо
                server.sendmail(sender_email, recipient_email, msg.as_string())
                print(f"Письмо для {full_name} ({recipient_email}) успешно отправлено.")

    except Exception as e:
        print(f"Произошла ошибка: {e}")

# Пример использования:
send_pdf_emails(
    folder_path="./pdf_files",  # Папка с PDF файлами
    excel_path="./recipients.xlsx",  # Путь к Excel файлу
    smtp_server="smtp.gmail.com",  # SMTP сервер
    smtp_port=587,  # Порт сервера
    sender_email="your_email@gmail.com",  # Ваш email
    sender_password="your_password"  # Ваш пароль
)
