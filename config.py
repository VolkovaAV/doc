# Логин по умолчанию для окна входа. Пароль программа запрашивает
# при каждом запуске рассылки и нигде не сохраняет.
FROM_MAIL = "pay.incas@mail.ru"

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
