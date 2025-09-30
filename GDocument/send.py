import smtplib
import pandas as pd
import numpy as np
import time
from .generate import email, fname
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
        decoded = f.decode()
        # Формат строки: (<атрибуты>) "<разделитель>" "<имя папки>"
        # Пример: (\\HasNoChildren \\Sent) "/" "Sent"
        if "\\Sent" in decoded:
            # имя папки в кавычках в конце строки
            return decoded.split(' "/" ')[-1].strip('"')

    # fallback, если сервер не метит \Sent
    return "Sent"

def send_email(df, testing, params):

    msg = MIMEMultipart()                                     # Создаем сообщение
    msg["From"] = config.FROM_MAIL                                   # Добавляем адрес отправителя
    msg['To'] = df['email']                                       # Добавляем адрес получателя
    msg["Subject"] = Header(f'Оплата рег.взноса {params['EVENT_NAME']}', 'utf-8')        # Пишем тему сообщения
    msg["Date"] = formatdate(localtime=True)                  # Дата сообщения
    msg.attach(MIMEText(email(df), 'html', 'utf-8'))  # Добавляем форматированный текст сообщения
    # Добавляем файл
    part = MIMEBase('application', "octet-stream")            # Создаем объект для загрузки файла
    part.set_payload(open('./files/pdf/'+fname(df, type='bill')+'.pdf',"rb").read())              # Подключаем файл
    encoders.encode_base64(part)
    
    
    part.add_header('Content-Disposition',
                    f'attachment; filename="{fname(df, type='bill')+'.pdf'}"')
    msg.attach(part)                                          # Добавляем файл в письмо

    

    part = MIMEBase('application', "octet-stream")            # Создаем объект для загрузки файла
    part.set_payload(open('./files/pdf/'+fname(df, type='act')+'.pdf',"rb").read())              # Подключаем файл
    encoders.encode_base64(part)
    
    
    part.add_header('Content-Disposition',
                    f'attachment; filename="{fname(df, type='act')}.pdf"')
    msg.attach(part)
    
    try:
        smtp = smtplib.SMTP(config.SERVER_ADR, 25)                       # Создаем объект для отправки сообщения 
        smtp.starttls()                                           # Открываем соединение
        smtp.ehlo()
        smtp.login(config.FROM_MAIL, config.FROM_PASSW)                        # Логинимся в свой ящик
        if testing:
            smtp.sendmail(config.FROM_MAIL, config.TO_MAIL_TEST, msg.as_string())
        else:
            smtp.sendmail(config.FROM_MAIL, df['email'], msg.as_string())
        smtp.quit()

        imap = imaplib.IMAP4_SSL(config.IMAP_SERVER, 993)                     # Подключаемся в почтовому серверу
        imap.login(config.FROM_MAIL, config.FROM_PASSW)                        # Логинимся в свой ящик
        path = find_sent_folder(imap)
        imap.select(path)                                       # Переходим в папку Исходящие
        imap.append(path, None,                                 # Добавляем наше письмо в папку Исходящие
                    imaplib.Time2Internaldate(time.time()),
                    msg.as_bytes())
        
        return 'Письмо отправлено:' + df['email'] + '\n'
    except:
        return df
    
def send_all(testing):
    params = load_config()
    xl = pd.read_excel(config.TB_NAME, dtype='str')
    df = pd.DataFrame(xl)

    df =df.rename(columns={'Фамилия': 'LAST_NAME', 'Имя': 'FIRST_NAME', 'Отчество': 'MIDDLE_NAME', 'Сумма': 'SUMM'})
    
    df['SEX'] = np.where(df['MIDDLE_NAME'].str.endswith('на'), 'ая',np.where(df['MIDDLE_NAME'].str.endswith('ич'), 'ый','ый(ая)')
)

    df['F_NAME'] = df['FIRST_NAME'].str[0] + '.'
    df['M_NAME'] = df['MIDDLE_NAME'].str[0] + '.'

    for person_ID in range(len(df)):
        res = ''
        # return 1
        send_email(df.iloc[person_ID], testing, params)
    return "Отправка завершена!"


