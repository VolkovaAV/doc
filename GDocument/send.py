import os
import smtplib
import time
from datetime import datetime

import pandas as pd
from .generate import email, fname, load_participants, PDF_DIR
from .json_work import *
import config

from email.mime.text import MIMEText
from email.header    import Header
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formatdate


def _smtp_text(code_resp):
    """(550, b'mailbox unavailable') -> '550 mailbox unavailable'"""
    code, resp = code_resp
    if isinstance(resp, bytes):
        resp = resp.decode('utf-8', errors='replace')
    return f"{code} {resp}"

def _error_text(e: smtplib.SMTPException) -> str:
    """Понятный текст ошибки SMTP для истории и отчета."""
    if isinstance(e, smtplib.SMTPRecipientsRefused):
        return '; '.join(_smtp_text(v) for v in e.recipients.values())
    if isinstance(e, smtplib.SMTPResponseException):
        return _smtp_text((e.smtp_code, e.smtp_error))
    return str(e)

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

REPORT_DIR = os.path.join(config.FILES_FOLDER_NAME, 'reports')

def save_report(df, results, testing):
    """
    Сохраняет итоги рассылки в новую Excel-таблицу files/reports/рассылка_<дата_время>.xlsx.
    results — список словарей (по одному на участника) со статусом отправки.
    Возвращает путь к файлу.
    """
    os.makedirs(REPORT_DIR, exist_ok=True)
    report = pd.DataFrame({
        'Фамилия': df['LAST_NAME'],
        'Имя': df['FIRST_NAME'],
        'Отчество': df['MIDDLE_NAME'].fillna(''),
        'email участника': df['email'],
        'email в таблице': df['EMAIL_RAW'],
        'Сумма': df['SUMM'].astype(int),
        'Счет': [fname(df.iloc[i], 'bill') + '.pdf' for i in range(len(df))],
        'Режим': 'тест' if testing else 'рассылка',
        'Отправлено на': [r['to'] for r in results],
        'Статус': [r['status'] for r in results],
        'Время отправки': [r['time'] for r in results],
        'Ошибка': [r['error'] for r in results],
    })

    kind = 'тест' if testing else 'рассылка'
    path = os.path.join(REPORT_DIR, f"{kind}_{datetime.now():%Y-%m-%d_%H-%M-%S}.xlsx")
    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        report.to_excel(writer, sheet_name='Отчет', index=False)
        sheet = writer.sheets['Отчет']
        sheet.freeze_panes = 'A2'                                  # шапка всегда видна
        sheet.auto_filter.ref = sheet.dimensions                   # фильтры по столбцам
        for column in sheet.columns:                               # ширина по содержимому
            width = max(len(str(c.value or '')) for c in column)
            sheet.column_dimensions[column[0].column_letter].width = min(width + 2, 60)
    return path

def send_all(testing, login, password):
    """
    Рассылает счета всем участникам из таблицы.
    login — адрес, с которого уходят письма (основной адрес или псевдоним ящика),
    password — пароль для внешних приложений основного ящика.
    Копия каждого письма отправляется на адрес login.
    testing=True — все письма уходят только на адрес login (письмо самому себе).
    Итоги сохраняются в новую Excel-таблицу (см. save_report).
    """
    if not login or not password:
        raise ValueError("Не указан логин или пароль почты.")

    params = load_config()
    require_event_params(params)
    df = load_participants()

    # Сначала собираем все письма: если какого-то PDF нет, не отправляем ничего
    messages = [build_message(df.iloc[person_ID], testing, params, login) for person_ID in range(len(df))]
    # статус по каждому участнику — для отчета
    results = [{'to': to_addr, 'status': 'Не отправлено', 'time': '', 'error': ''} for _, to_addr in messages]

    sent, failed = [], []
    report_lines = []
    smtp = smtplib.SMTP_SSL(config.SERVER_ADR, config.SMTP_PORT, timeout=60)
    try:
        try:
            smtp.login(login, password)                       # Логинимся в ящик (можно под псевдонимом)
        except smtplib.SMTPAuthenticationError as e:
            raise RuntimeError("Неверный логин или пароль почты (для mail.ru нужен "
                               f"пароль для внешних приложений). Письма не отправлены. {e}") from None

        aborted = None  # причина аварийной остановки рассылки
        for (msg, to_addr), result in zip(messages, results):
            # Копия письма уходит отправителю как скрытая копия (участник ее не видит).
            # В тестовом режиме to_addr уже равен login — второе письмо не нужно.
            recipients = [to_addr] if to_addr.lower() == login.lower() else [to_addr, login]
            try:
                refused = smtp.send_message(msg, from_addr=login, to_addrs=recipients)
            except smtplib.SMTPException as e:
                error = _error_text(e)
                result.update(status='Ошибка', error=error)
                if _sender_rejected(e):
                    # сервер не разрешает такого отправителя — остальные письма тоже не уйдут
                    aborted = f"Сервер отклонил отправителя {login}. Отправлено писем: {len(sent)}. {error}"
                    break
                failed.append(f"{to_addr}: {error}")
                continue

            # сервер мог принять письмо только для части адресов
            if to_addr in refused:
                error = _smtp_text(refused[to_addr])
                result.update(status='Ошибка', error=error)
                failed.append(f"{to_addr}: {error}")
                continue
            note = 'копия отправителю не доставлена' if login in refused else ''
            result.update(status='Отправлено', time=f"{datetime.now():%d.%m.%Y %H:%M:%S}", error=note)
            sent.append(f"Письмо отправлено: {to_addr}" + (f" ({note})" if note else ''))
            time.sleep(config.SEND_DELAY)

        # отчет сохраняем и тогда, когда рассылка прервалась на середине
        try:
            report_lines.append(f"Отчет о рассылке: {save_report(df, results, testing)}")
        except Exception as e:
            report_lines.append(f"ВНИМАНИЕ: не удалось сохранить отчет о рассылке: {e}")
        if aborted:
            raise RuntimeError('\n'.join([aborted] + report_lines))
    finally:
        try:
            smtp.quit()
        except smtplib.SMTPException:
            pass

    lines = df.attrs.get('notes', []) + sent + [f"ОШИБКА отправки {x}" for x in failed]
    lines.append(f"Отправка завершена: отправлено {len(sent)}, ошибок {len(failed)}.")
    lines += report_lines
    return '\n'.join(lines)
