# Файл main #

- В файле "main" скрипт под файлы формата "png". Нужна таблица exel, в которой 2 столбца с ФИО:Почта.
- `process_emails` функция которая отправляет только тем, у кого есть файл в папке с ФИО (Поиск идет по наименованию файла)

# Пример использования:
- В самом конце кода, надо указать почту с которой будет отправка и пути к файлу exel и папке с файлами.
```python
excel_file = 'users.xlsx'  # Путь к Excel файлу
images_folder = 'images'  # Путь к папке с изображениями
sender_email = 'your_email@gmail.com'  # Ваш email
sender_password = 'your_email_password'  # Ваш пароль или token
```

# PDF

- В файле "PDF" скрипт под файлы формата "pdf".
- `send_pdf_emails` Рассылает PDF файлы пользователям на основе Excel таблицы.
```python
"""
folder_path: Путь к папке с PDF файлами
excel_path: Путь к Excel файлу с колонками 'ФИО' и 'Email'
smtp_server: Адрес SMTP сервера
smtp_port: Порт SMTP сервера
sender_email: Email отправителя
sender_password: Пароль от email отправителя
"""
```

# PDF2.0_API

- В файле "PDF2.0" скрипт такой же как и в "PDF". 
- `send_pdf_emails_via_api` Рассылает PDF файлы пользователям на основе Excel таблицы через API.
```python
"""
folder_path: Путь к папке с PDF файлами
excel_path: Путь к Excel файлу с колонками 'ФИО' и 'Email'
token: Токен доступа для API (OAuth2)
"""
```
- Важно указать базовый URL для отправки писем через Microsoft Graph API.
```
base_url = "https://graph.microsoft.com/v1.0/me/sendMail"
```
- `send_pdf_emails_via_api` в этой части указываем пути и токен.
# Что нужно настроить:
- Настройка для файла PDF2.0_API
- Получите OAuth2 токен от Microsoft Graph API, связавшись с администратором вашей почтовой системы. 
- Убедитесь, что разрешения API включают отправку писем (Mail.Send).