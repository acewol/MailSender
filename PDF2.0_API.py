import os
import pandas as pd
import requests
from requests.structures import CaseInsensitiveDict

def send_pdf_emails_via_api(folder_path, excel_path, token):
    """
    Рассылает PDF файлы пользователям на основе Excel таблицы через API.

    :param folder_path: Путь к папке с PDF файлами
    :param excel_path: Путь к Excel файлу с колонками 'ФИО' и 'Email'
    :param token: Токен доступа для API (OAuth2)
    """
    try:
        # Загружаем Excel файл
        data = pd.read_excel(excel_path)

        # Проверяем наличие необходимых колонок
        if not {'ФИО', 'Email'}.issubset(data.columns):
            raise ValueError("Excel файл должен содержать колонки 'ФИО' и 'Email'")

        # Базовый URL для отправки писем через Microsoft Graph API
        base_url = "https://graph.microsoft.com/v1.0/me/sendMail"

        headers = CaseInsensitiveDict()
        headers["Authorization"] = f"Bearer {token}"
        headers["Content-Type"] = "application/json"

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

            # Читаем содержимое файла
            with open(pdf_file_path, "rb") as f:
                pdf_content = f.read()

            # Кодируем файл в Base64
            encoded_pdf = pdf_content.encode('base64').decode('utf-8')

            # Формируем тело запроса
            email_body = {
                "message": {
                    "subject": "Ваш документ",
                    "body": {
                        "contentType": "Text",
                        "content": f"Уважаемый(ая) {full_name},\n\nВо вложении вы найдете ваш документ."
                    },
                    "toRecipients": [
                        {"emailAddress": {"address": recipient_email}}
                    ],
                    "attachments": [
                        {
                            "@odata.type": "#microsoft.graph.fileAttachment",
                            "name": pdf_file_name,
                            "contentBytes": encoded_pdf
                        }
                    ]
                }
            }

            # Отправляем запрос
            response = requests.post(base_url, headers=headers, json=email_body)

            if response.status_code == 202:
                print(f"Письмо для {full_name} ({recipient_email}) успешно отправлено.")
            else:
                print(f"Не удалось отправить письмо для {full_name} ({recipient_email}). Ошибка: {response.status_code}, {response.text}")

    except Exception as e:
        print(f"Произошла ошибка: {e}")

# Пример использования:
send_pdf_emails_via_api(
    folder_path="./pdf_files",  # Папка с PDF файлами
    excel_path="./recipients.xlsx",  # Путь к Excel файлу
    token="your_access_token"  # Токен доступа OAuth2
)
