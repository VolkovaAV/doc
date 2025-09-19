import smtplib
import pandas as pd
from .generate import email, fname


from email.mime.text import MIMEText
from email.header    import Header
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formatdate

from_mail = "pay.incas@mail.ru"                         # Почта отправителя
from_passwd = "8sRt0H0xDvqTmCxhNGx8"                      # пароль от почты отправителя
server_adr = "smtp.mail.ru"                               # адрес почтового сервера

to_mail_test = 'a.volkova@ipfran.ru'                           # адрес получателя
# to_mail_test = 'volkovaav2017@gmail.com'                           # адрес получателя


def send_email(df, testing= True, first=True):
    global from_mail
    global from_passwd
    global server_adr

    msg = MIMEMultipart()                                     # Создаем сообщение
    msg["From"] = from_mail                                   # Добавляем адрес отправителя
    msg['To'] = df['email']                                       # Добавляем адрес получателя
    msg["Subject"] = Header('Оплата рег.взноса NWP-2025', 'utf-8')        # Пишем тему сообщения
    msg["Date"] = formatdate(localtime=True)                  # Дата сообщения
    msg.attach(MIMEText(email(df, first), 'html', 'utf-8'))  # Добавляем форматированный текст сообщения
    # Добавляем файл
    part = MIMEBase('application', "octet-stream")            # Создаем объект для загрузки файла
    part.set_payload(open('./files/out/'+fname(df, type='bill')+'.pdf',"rb").read())              # Подключаем файл
    encoders.encode_base64(part)
    
    
    part.add_header('Content-Disposition',
                    f'attachment; filename="{fname(df, type='bill')+'.pdf'}"')
    msg.attach(part)                                          # Добавляем файл в письмо

    

    part = MIMEBase('application', "octet-stream")            # Создаем объект для загрузки файла
    part.set_payload(open('./files/out/'+fname(df, type='act')+'.pdf',"rb").read())              # Подключаем файл
    encoders.encode_base64(part)
    
    
    part.add_header('Content-Disposition',
                    f'attachment; filename="{fname(df, type='act')}.pdf"')
    msg.attach(part)
    
    try:
        smtp = smtplib.SMTP(server_adr, 25)                       # Создаем объект для отправки сообщения 
        smtp.starttls()                                           # Открываем соединение
        smtp.ehlo()
        smtp.login(from_mail, from_passwd)                        # Логинимся в свой ящик
        if not testing:
            smtp.sendmail(from_mail, df['email'], msg.as_string())
            print('Sent to ', df['email'])
            # smtp.sendmail(from_mail, 'aries@ipfran.ru', msg.as_string())
            smtp.sendmail(from_mail, 'a_evtushenko@inbox.ru', msg.as_string())        # Отправляем сообщения
        smtp.sendmail(from_mail, 'a.volkova@ipfran.ru', msg.as_string())
        

        smtp.quit()   
        print('Sent')
        return 0      
    except:
        return df

