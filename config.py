import os

FROM_MAIL = "pay.incas@mail.ru"                          # Почта отправителя
# Пароль НЕ хранится в репозитории: задайте его в config_local.py
# (см. config_local.example.py) или в переменной окружения DOCAPP_MAIL_PASSWORD
FROM_PASSW = os.getenv("DOCAPP_MAIL_PASSWORD", "")

SERVER_ADR = "smtp.mail.ru"                               # адрес почтового сервера (SMTP)
SMTP_PORT = 465                                           # SMTP over SSL
IMAP_SERVER = "imap.mail.ru"
IMAP_PORT = 993

TO_MAIL_TEST = 'a_evtushenko@inbox.ru'                    # адрес для тестовой рассылки
SEND_DELAY = 0.5                                          # пауза между письмами, сек


NameOrg = 'МЦФПИН'
PersonalAcc = '40703810942000000672'
BankName = 'ВОЛГО-ВЯТСКИЙ БАНК ПАО СБЕРБАНК'
BIC = '042202603'
CorrespAcc = '30101810900000000603'
KPP = '526001001'
PayeeINN = '5260054053'

HeadPosition = 'Ректор'                                   # должность подписанта
HeadName = 'А.А.Евтушенко'                                # подпись в счете

TEMP_FOLDER_NAME = 'templates'
FILES_FOLDER_NAME = 'files'

STD_COL_NAME = ["Фамилия", "Имя", "Отчество", "email", "Сумма"]
TB_NAME = "participant_list.xlsx"

# Локальные переопределения (пароль и т.п.), файл не хранится в git
try:
    from config_local import *  # noqa: F401,F403
except ImportError:
    pass
