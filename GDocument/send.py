import smtplib
import pandas as pd
import numpy as np
from .generate import email, fname
from .json_work import *


from email.mime.text import MIMEText
from email.header    import Header
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formatdate

import _config


def send_email(df, testing, params):

    msg = MIMEMultipart()                                     # Создаем сообщение
    msg["From"] = _config.FROM_MAIL                                   # Добавляем адрес отправителя
    msg['To'] = df['email']                                       # Добавляем адрес получателя
    msg["Subject"] = Header(f'Оплата рег.взноса {params['generate_parameters']['EVENT_NAME']}', 'utf-8')        # Пишем тему сообщения
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
        smtp = smtplib.SMTP(_config.SERVER_ADR, 25)                       # Создаем объект для отправки сообщения 
        smtp.starttls()                                           # Открываем соединение
        smtp.ehlo()
        smtp.login(_config.FROM_MAIL, _config.FROM_PASSW)                        # Логинимся в свой ящик
        if not testing:
            # smtp.sendmail(_config.FROM_MAIL, df['email'], msg.as_string())
            print('Sent to ', df['email'])
        smtp.sendmail(_config.FROM_MAIL, _config.TO_MAIL_TEST, msg.as_string())
        

        smtp.quit()   
        print('Sent')
        return 'Письмо отправлено:' + df['email'] + '\n'
    except:
        return df
    
def send_all(testing):
    params = load_config_json()
    xl = pd.read_excel(params['generate_parameters']['TB_NAME'], dtype='str')
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


