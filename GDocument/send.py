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


def _sender_rejected(e: smtplib.SMTPException) -> bool:
    """
    Отказ из-за адреса отправителя (а не из-за конкретного получателя).
    mail.ru отвечает на это «501 sender address must match authenticated user»
    уже на этапе получателя, поэтому смотрим текст ответа сервера.
    """
    if isinstance(e, smtplib.SMTPSenderRefused):
        return True
    if isinstance(e, smtplib.SMTPRecipientsRefused):
        responses = [resp for _, resp in e.recipients.values()]
    elif isinstance(e, smtplib.SMTPResponseException):
        responses = [e.smtp_error]
    else:
        return False
    return any(b'sender' in (r if isinstance(r, bytes) else str(r).encode()).lower() for r in responses)

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
    msg["From"] = from_addr                                   # Адрес отправителя (основной или псевдоним)
    msg["Reply-To"] = from_addr                               # Ответы участников придут на этот же адрес
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

def _open_imap(login, password):
    """
    Подключение к IMAP для сохранения копий в «Отправленные».
    Не обязательно для рассылки: при ошибке возвращает (None, текст ошибки).
    """
    try:
        imap = imaplib.IMAP4_SSL(config.IMAP_SERVER, config.IMAP_PORT, timeout=60)
    except OSError as e:
        return None, str(e)
    try:
        imap.login(login, password)
        return imap, ''
    except imaplib.IMAP4.error as e:
        try:
            imap.logout()
        except (imaplib.IMAP4.error, OSError):
            pass
        return None, str(e)

def send_all(testing, login, password):
    """
    Рассылает счета всем участникам из таблицы.
    login — адрес, с которого уходят письма (основной адрес или псевдоним ящика),
    password — пароль для внешних приложений основного ящика.
    testing=True — все письма уходят на адрес login (письмо самому себе).
    """
    if not login or not password:
        raise ValueError("Не указан логин или пароль почты.")

    params = load_config()
    require_event_params(params)
    df = load_participants()

    # Сначала собираем все письма: если какого-то PDF нет, не отправляем ничего
    messages = [build_message(df.iloc[person_ID], testing, params, login) for person_ID in range(len(df))]

    sent, failed, warnings = [], [], []
    smtp = smtplib.SMTP_SSL(config.SERVER_ADR, config.SMTP_PORT, timeout=60)
    imap = None
    try:
        try:
            smtp.login(login, password)                       # Логинимся в ящик (можно под псевдонимом)
        except smtplib.SMTPAuthenticationError as e:
            raise RuntimeError("Неверный логин или пароль почты (для mail.ru нужен "
                               f"пароль для внешних приложений). Письма не отправлены. {e}") from None

        imap, imap_error = _open_imap(login, password)
        if imap is None:
            warnings.append("ВНИМАНИЕ: не удалось войти по IMAP, копии писем не сохранены "
                            f"в «Отправленные». {imap_error}")
        else:
            sent_folder = find_sent_folder(imap)

        for msg, to_addr in messages:
            try:
                smtp.send_message(msg, from_addr=login, to_addrs=[to_addr])
            except smtplib.SMTPException as e:
                if _sender_rejected(e):
                    # сервер не разрешает такого отправителя — остальные письма тоже не уйдут
                    raise RuntimeError(
                        f"Сервер отклонил отправителя {login}. Отправлено писем: {len(sent)}. {e}") from None
                failed.append(f"{to_addr}: {e}")
                continue

            note = ''
            if imap is not None:
                # Кладем копию письма в папку «Отправленные»
                status, _ = imap.append(sent_folder, '\\Seen',
                                        imaplib.Time2Internaldate(time.time()),
                                        msg.as_bytes())
                if status != 'OK':
                    note = ' (не удалось сохранить в «Отправленные»)'
            sent.append(f"Письмо отправлено: {to_addr}{note}")
            time.sleep(config.SEND_DELAY)
    finally:
        try:
            smtp.quit()
        except smtplib.SMTPException:
            pass
        if imap is not None:
            try:
                imap.logout()
            except imaplib.IMAP4.error:
                pass

    lines = warnings + sent + [f"ОШИБКА отправки {x}" for x in failed]
    lines.append(f"Отправка завершена: отправлено {len(sent)}, ошибок {len(failed)}.")
    return '\n'.join(lines)
