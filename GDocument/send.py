import os
import re
import smtplib
import time
from .generate import email, fname, load_participants, PDF_DIR
from .json_work import *
import config
import imaplib

from email.mime.text import MIMEText
from email.header    import Header
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formatdate


def find_sent_folder(imap: imaplib.IMAP4_SSL) -> str:
    """
    Возвращает название папки "Отправленные" для текущего IMAP-сервера.
    Если не удалось найти — возвращает 'Sent'.
    """
    status, folders = imap.list()
    if status != "OK":
        return "Sent"

    for f in folders:
        decoded = f.decode(errors="replace")
        # Формат строки: (<атрибуты>) "<разделитель>" <имя папки>
        # Пример: (\HasNoChildren \Sent) "/" "Sent"  или  (\Sent) "|" Sent
        if "\\Sent" in decoded:
            m = re.match(r'\((?P<attrs>[^)]*)\)\s+(?:"[^"]*"|NIL)\s+(?P<name>.+)$', decoded)
            if m:
                name = m.group("name").strip()
                # имя папки передаём серверу в кавычках, как оно пришло
                return name if name.startswith('"') else f'"{name}"'

    # fallback, если сервер не метит \Sent
    return "Sent"

def build_message(df, testing, params, from_addr):
    """Собирает письмо со счетом для одного участника."""
    to_addr = from_addr if testing else df['email']  # тест — письмо самому себе

    msg = MIMEMultipart()                                     # Создаем сообщение
    msg["From"] = from_addr                                   # Добавляем адрес отправителя
    msg['To'] = to_addr                                       # Добавляем адрес получателя
    msg["Subject"] = Header(f"Оплата рег.взноса {params['EVENT_NAME']}", 'utf-8')  # Пишем тему сообщения
    msg["Date"] = formatdate(localtime=True)                  # Дата сообщения
    msg.attach(MIMEText(email(df), 'html', 'utf-8'))          # Добавляем форматированный текст сообщения

    # Добавляем файл
    bill_name = fname(df, type='bill') + '.pdf'
    part = MIMEBase('application', "pdf")                     # Создаем объект для загрузки файла
    with open(os.path.join(PDF_DIR, bill_name), "rb") as f:
        part.set_payload(f.read())                            # Подключаем файл
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment', filename=bill_name)
    msg.attach(part)                                          # Добавляем файл в письмо

    return msg, to_addr

def send_all(testing, login, password):
    """
    Рассылает счета всем участникам из таблицы.
    testing=True — все письма уходят на адрес отправителя (login).
    login, password — учетные данные почты-отправителя (вводятся в окне входа).
    """
    if not login or not password:
        raise ValueError("Не указан логин или пароль почты.")

    params = load_config()
    require_event_params(params)
    df = load_participants()

    # Сначала собираем все письма: если какого-то PDF нет, не отправляем ничего
    messages = [build_message(df.iloc[person_ID], testing, params, login) for person_ID in range(len(df))]

    sent, failed = [], []
    smtp = smtplib.SMTP_SSL(config.SERVER_ADR, config.SMTP_PORT, timeout=60)
    imap = imaplib.IMAP4_SSL(config.IMAP_SERVER, config.IMAP_PORT, timeout=60)
    try:
        try:
            smtp.login(login, password)                       # Логинимся в свой ящик
            imap.login(login, password)
        except (smtplib.SMTPAuthenticationError, imaplib.IMAP4.error) as e:
            raise RuntimeError("Неверный логин или пароль почты (для mail.ru нужен "
                               f"пароль для внешних приложений). Письма не отправлены. {e}") from None
        sent_folder = find_sent_folder(imap)

        for msg, to_addr in messages:
            try:
                smtp.sendmail(login, to_addr, msg.as_string())
            except smtplib.SMTPException as e:
                failed.append(f"{to_addr}: {e}")
                continue

            # Кладем копию письма в папку «Отправленные»
            status, _ = imap.append(sent_folder, '\\Seen',
                                    imaplib.Time2Internaldate(time.time()),
                                    msg.as_bytes())
            note = '' if status == 'OK' else ' (не удалось сохранить в «Отправленные»)'
            sent.append(f"Письмо отправлено: {to_addr}{note}")
            time.sleep(config.SEND_DELAY)
    finally:
        try:
            smtp.quit()
        except smtplib.SMTPException:
            pass
        try:
            imap.logout()
        except imaplib.IMAP4.error:
            pass

    lines = sent + [f"ОШИБКА отправки {x}" for x in failed]
    lines.append(f"Отправка завершена: отправлено {len(sent)}, ошибок {len(failed)}.")
    return '\n'.join(lines)
