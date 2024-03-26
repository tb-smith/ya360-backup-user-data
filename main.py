from imapclient import IMAPClient, AbortError
import email
import time
import os
import logging
from email.header import decode_header
import requests
import shutil
from dotenv import load_dotenv
import yadisk
import shutil


def archive_folder(archive_name, folder_path):
    shutil.make_archive(archive_name, 'zip', folder_path)
    shutil.rmtree(folder_path)
    print(f"Папка {folder_path} архивирована и удалена.")
    return True


def get_yandex_token(client_id, client_secret, user_email):
    url = "https://oauth.yandex.ru/token"
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    payload = {
        'grant_type': 'urn:ietf:params:oauth:grant-type:token-exchange',
        'client_id': client_id,
        'client_secret': client_secret,
        'subject_token': user_email,
        'subject_token_type': 'urn:yandex:params:oauth:token-type:email'
    }
    
    response = requests.post(url, headers=headers, data=payload)
    
    if response.status_code == 200:
        # Успешно получили токен, возвращаем его
        return response.json().get('access_token')
    else:
        # Что-то пошло не так
        print(f"Ошибка: {response.status_code}, {response.json().get('error_description', 'unknown error')}")
        return None


def save_email(folder, msg):
    subject, encoding = decode_header(msg["Subject"])[0] if msg["Subject"] else (b"Unknown Subject", None)
    try:
        if isinstance(subject, bytes):
            subject = subject.decode(encoding if encoding else "utf-8")
    except LookupError:
        subject = subject.decode('utf-8', errors='replace')  # Использование 'utf-8' и замена нераспознанных символов

    subject = subject.replace('/', '-').replace('\\', '-')  # Удаление недопустимых символов из имени файла
    max_length = 100  # максимальная длина имени файла
    subject = subject[:max_length]  # обрезаем до 100 символов

    with open(f"{folder}/{subject}.eml", "wb") as f:
        f.write(msg.as_bytes())
        print(f"Сохранил письмо: {subject}")


def delete_folder(folder_path):
    shutil.rmtree(folder_path)
    print(f"Папка {folder_path} удалена")


# Функция для рекурсивной загрузки файлов и папок
def download_directory(y, path, local_path):
    os.makedirs(local_path, exist_ok=True)
    for item in y.listdir(path):
        if item.type == 'dir':
            download_directory(y, item.path, os.path.join(local_path, item.name))
        else:
            y.download(item.path, os.path.join(local_path, item.name))


if __name__ == "__main__":

    load_dotenv()

    logging.basicConfig(
        level=logging.INFO,  # Общий уровень логирования
        format="%(asctime)s [%(levelname)s]: %(message)s",
    )

    # Обработчик для вывода в файл с уровнем INFO
    file_handler = logging.FileHandler("mail_sync.log")
    file_handler.setLevel(logging.INFO)

    # Обработчик для вывода в консоль с уровнем WARNING
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.WARNING)

    # Добавление обработчиков
    logger = logging.getLogger()
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    client_id_disk = os.getenv("CLIENT_ID_DISK")
    client_secret_disk = os.getenv("CLIENT_SECRET_DISK")
    target_dir = os.getenv("TARGET_DIR")
    client_id_mail = os.getenv("CLIENT_ID_MAIL")
    client_secret_mail = os.getenv("CLIENT_SECRET_MAIL")
    target_dir = os.getenv("TARGET_DIR")
    users = os.getenv("USERS")
    start_time = time.time()  # Засекаем начальное время

    with open('first.txt', 'r') as f:
        users = f.read().splitlines()
        for user_email in users:
            downloaded_emails = 0  # счетчик успешно скачанных писем
            last_processed = 0  # счетчик последнего обработанного письма
            # подключение и авторизация
            # Архивирую почту
            user_name = user_email.split('@')[0]
            share_folder_user_path = f"{target_dir}/{user_name}"
            share_folder_user_mail_path = f"{target_dir}/{user_name}/mail"
            user_mail_archive_name = f"{target_dir}/{user_name}/{user_name}-mail"
            user_token = get_yandex_token(client_id_mail, client_secret_mail, user_email)
            logging.info(f"Подключение к почтовому ящику {user_email}")
            with IMAPClient("imap.yandex.ru") as client:
                client.oauth2_login(user_email, user_token, mech='XOAUTH2')
                # получение списка папок и создание их на диске
                total_emails = 0  # счетчик общего количества писем
                folders = client.list_folders()
                for folder in folders:
                    folder_path = os.path.join(share_folder_user_mail_path, folder[-1])
                    if not os.path.exists(folder_path):
                        os.makedirs(folder_path)
                    client.select_folder(folder[-1], readonly=True)
                    email_ids = client.search()
                    total_emails += len(email_ids) # увеличиваем счетчик на количество писем в папке
                    logging.info(f"Папка {folder[-1]}: {total_emails} писем")
                    chunk = []  # Определим переменную до цикла
                    chunk_size = 10
                    while last_processed < len(email_ids):
                        chunk = email_ids[last_processed:last_processed + chunk_size]
                        try:
                            response = client.fetch(chunk, ['BODY[]'])
                            for msg_id, data in response.items():
                                if b'BODY[]' not in data:
                                    logging.warning(f"Не удалось получить тело письма с ID {msg_id} {data}. Пропускаем.")
                                    continue  # пропустить текущую итерацию и перейти к следующему сообщению
                                raw_email = data[b'BODY[]']
                                msg = email.message_from_bytes(raw_email)
                                save_email(os.path.join(share_folder_user_mail_path, folder[-1]), msg)
                                last_processed += 1  # увеличиваем счетчик последнего обработанного письма на 1
                                downloaded_emails += 1  # увеличиваем счетчик на 1   
                                print(f"Скачано {downloaded_emails} писем из {total_emails}", end="\r")
                        except AbortError as e:
                            logging.error(f"Произошла ошибка: {e}")
                            time.sleep(10)  # Пауза перед повторным соединением
                            # Код для повторного соединения...
                            user_token = get_yandex_token(client_id_mail, client_secret_mail, user_email)
                            client.oauth2_login(user_email, user_token, mech='XOAUTH2')
                            client.select_folder(folder[-1], readonly=True)

                shutil.make_archive(user_mail_archive_name, 'zip', share_folder_user_mail_path)
                logging.info(f"Архив {user_mail_archive_name} создан")
                shutil.rmtree(share_folder_user_mail_path)
                logging.info(f"Папка {share_folder_user_mail_path} удалена")

                end_time = time.time()  # Засекаем конечное время
                elapsed_time = end_time - start_time  # Вычисляем затраченное время в секундах
                elapsed_minutes = elapsed_time / 60  # Преобразуем в минуты

                logging.info(f"Загрузка завершена! Время выполнения: {elapsed_minutes:.2f} минут")
                logging.info(f"Загружено {downloaded_emails} писем")
                logging.info(f"Всего писем: {total_emails}")
            # Архивирую диск
            user_token = get_yandex_token(client_id_disk, client_secret_disk, user_email)
            y = yadisk.YaDisk(client_id_disk, client_secret_disk, user_token)
            # Получаем информацию о диске
            user_used_space = y.get_disk_info()['used_space']
            trash_used_space = y.get_disk_info()['trash_size']
            total_space_for_backup = user_used_space + trash_used_space
            total_size_mb = total_space_for_backup / (2**20)
            print(f"Общий размер: {total_size_mb:.2f} Мб") # Выводим размер в Мб
            # Загрузка содержимого Яндекс.Диска
            root_path = '/' # корневая папка Яндекс.Диска
            share_folder_user_path = f"{target_dir}/{user_name}"
            share_folder_user_disk_path = f"{target_dir}/{user_name}/disk" # локальная папка для сохранения диска
            user_disk_archive_name = f"{target_dir}/{user_name}/{user_name}-disk"

            if trash_used_space > 0:
                print("В корзине есть файлы, восстанавливаю")
                for item in y.trash_listdir("/"):
                    y.restore_trash(item.path)
            print("Загружаю диск")
            download_directory(y, root_path, share_folder_user_disk_path)

            # Архивирование загруженных файлов
            print("Архивирую папку диска")

            shutil.make_archive(user_disk_archive_name, 'zip', share_folder_user_disk_path)
            shutil.rmtree(share_folder_user_disk_path)
            print(f"Папка {share_folder_user_disk_path} архивирована и удалена.")
            print("Завершено")
