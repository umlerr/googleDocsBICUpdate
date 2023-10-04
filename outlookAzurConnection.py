import os
import datetime
from dotenv import load_dotenv
from O365 import Account, FileSystemTokenBackend

# Загрузите конфигурационный файл, содержащий данные о подключении к API Outlook
load_dotenv('outlook.env')

# Получите значения переменных окружения из файла .env
client_id = os.getenv('CLIENT_ID')
client_secret = os.getenv('CLIENT_SECRET')
token_file_path = os.getenv('TOKEN_FILE_PATH')

# Создайте строку для авторизации
credentials = (client_id, client_secret)

# Создайте экземпляр объекта Account
account = Account(credentials)#, token_backend=FileSystemTokenBackend(token_path=token_file_path))

# Аутентифицируйте пользователя
if not account.is_authenticated:
    account.authenticate(scopes=['basic', 'message_all'])

# Получите папку "Входящие"
inbox_folder = account.inbox_folder()

# Определите отправителя и тему фильтра
sender_email = 'outlook_5e612556ed318093@outlook.com'  #'kerrychen@contcircle.com'
email_subject_keywords = 'Shanghai, Inventory, daily report'

# Настройте параметры поиска
search_criteria = inbox_folder.new_query()
search_criteria = search_criteria.on_attribute('sender').equals(sender_email)
search_criteria = search_criteria.chain('AND').on_attribute('subject').contains(email_subject_keywords)

# Выполните поиск сообщений с заданными критериями
messages = inbox_folder.get_messages(limit=50, query=search_criteria)

# Работайте с найденными сообщениями
for message in messages:
    # Работайте с прикрепленными файлами
    attachments = message.get_attachments()
    for attachment in attachments:
        if attachment.is_file:
            # Задайте путь и имя файла для сохранения
            today = datetime.datetime.now().strftime("%Y-%m-%d")
            save_path = f'F:/Work/{today}/'
            os.makedirs(save_path, exist_ok=True)
            file_path = os.path.join(save_path, attachment.name)
            # Сохраните файл
            with open(file_path, 'wb') as f:
                f.write(attachment.content)
                print(f'DailyReport: {attachment.name} to {file_path}')

# Завершаем работу программы
account.logout()