import os
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.utils import formataddr

def send_email(sender_email, sender_password, recipient_email, subject, body, attachment=None):
    """
    Функция для отправки email с возможностью добавления вложений.
    """
    # Настройка SMTP-сервера
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(sender_email, sender_password)

    # Создание письма
    msg = MIMEMultipart()
    msg['From'] = formataddr(('Sender', sender_email))
    msg['To'] = recipient_email
    msg['Subject'] = subject

    # Добавление текста письма
    msg.attach(MIMEText(body, 'plain'))

    # Добавление вложений (если есть)
    if attachment:
        with open(attachment, 'rb') as f:
            img = MIMEImage(f.read())
            msg.attach(img)

    # Отправка письма
    server.sendmail(sender_email, recipient_email, msg.as_string())
    server.quit()

def process_emails(excel_file, images_folder, sender_email, sender_password):
    """
    Основная функция для отправки email только тем пользователям, у которых есть файл в папке.
    """


    # Чтение Excel-файла
    df = pd.read_excel(excel_file)

    # Проходим по всем строкам в DataFrame
    for index, row in df.iterrows():
        user_name = row['ФИО']
        user_email = row['Email']
        # Формируем имя файла из ФИО (например, 'Иванов Иван Иванович.png')
        image_filename = f"{user_name}.png"
        image_path = os.path.join(images_folder, image_filename)

        # Проверяем, есть ли файл в папке
        if os.path.exists(image_path):
            # Если файл существует, отправляем письмо
            subject = f"Привет, {user_name}!"
            body = f"Здравствуйте, {user_name}!\n\nВам отправляется файл, связанный с вашим именем."

            send_email(sender_email, sender_password, user_email, subject, body, attachment=image_path)
            print(f"Email отправлен пользователю: {user_name} ({user_email})")

        else:
            print(f"Файл для {user_name} не найден. Пропускаем рассылку.")

# Пример использования:
# Должен быть путь к Excel файлу и к папке с изображениями.
excel_file = 'users.xlsx'  # Путь к Excel файлу
images_folder = 'images'  # Путь к папке с изображениями
sender_email = 'your_email@gmail.com'  # Ваш email
sender_password = 'your_email_password'  # Ваш пароль или token

process_emails(excel_file, images_folder, sender_email, sender_password)
