import qrcode # для генерации кода
import codecs # для обработки рускоязычных файлов

import pandas as pd

import subprocess # для запуска процессов s



NameOrg = 'МЦФПИН'
PersonalAcc = '40703810942000000672'
BankName= 'ВОЛГО-ВЯТСКИЙ БАНК ПАО СБЕРБАНК'
BIC = '042202603'
CorrespAcc='30101810900000000603'
KPP='526001001'
PayeeINN='5260054053'

def qr_code(df):
    '''
    Функция сохраняет в папку проекта картинку с qr кодом для оплаты.
    Картинка сохраняется в папку по адресу: /files/qr_code
    Название файла генерируется при помощи функции fname() с параметром type='qr'

    Вход:
        df - pd.Series - информация об участнике
    Выход:
        path_name - название сохраненного файла
    '''
    global NameOrg
    global PersonalAcc
    global BankName
    global BIC
    global CorrespAcc
    global KPP
    global PayeeINN

    if df["MiddleName"]!=df["MiddleName"]:
        PersonInfo = '(участник '+df['LastName']+' '+df['FirstName'][0]+'.)'
    else:
        PersonInfo = '(участник '+df['LastName']+' '+df['FirstName'][0]+'. '+df['MiddleName'][0]+'.)'

    data=f'ST00012|'\
        f'Name={NameOrg}|'\
        f'PersonalAcc={PersonalAcc}|'\
        f'BankName={BankName}|'\
        f'BIC={BIC}|'\
        f'CorrespAcc={CorrespAcc}|'\
        f'KPP={KPP}|'\
        f'PayeeINN={PayeeINN}|'\
        f'Purpose=Услуги по организации питания 6-ти участников конференции VNM-2025|'\
        f'SUM={int(df['SUM'])*100}'
    
    img = qrcode.make(data)
    path_name = f'./files/qr_code/kopilova.png'
    img.save(path_name)
    return path_name

def bill(df):
    '''
    Функция генерирует tex-файл для счета по шаблону
    Шаблон счета расположен: /templates/bill.tex
    Путь к файлу: /files
    Название файла генерируется при помощи функции fname с параметром type='bill'
    Вход:
        df - pd.Series - информация об участнике
    Выход:
        Путь к сохраненному файлу 

    '''

    global NameOrg
    global PersonalAcc
    global BankName
    global BIC
    global CorrespAcc
    global KPP
    global PayeeINN

    fileObj = codecs.open( "templates/bill_new.tex", "r", "utf_8" )
    text = fileObj.read()

    text_tmp = text.replace('NAMEORG', NameOrg)
    text_tmp = text_tmp.replace('INN', PayeeINN)
    text_tmp = text_tmp.replace('NUMBERCHET', PersonalAcc)
    text_tmp = text_tmp.replace('NAMEBANK', BankName)
    text_tmp = text_tmp.replace('BICBANK', BIC)
    text_tmp = text_tmp.replace('KORRCHET', CorrespAcc)
    text_tmp = text_tmp.replace('KPP', KPP)

    # text_tmp = text_tmp.replace('DOCNUM', df['DOCNUM'])

    text_tmp = text_tmp.replace('PLATE', f'{df["LastName"]} {df["FirstName"]} {df["MiddleName"]}')
    filename = f'{df['LastName']}_{df["FirstName"][0]}_{df["MiddleName"][0]}_{df["SUM"]}'

    text_tmp1 = text_tmp.replace('SUM', df['SUM'])
    text_tmp1 = text_tmp1.replace('FILENAME.png', qr_code(df))

    f_temp = codecs.open(f'files/kopilova.tex', 'w', "utf_8")
    f_temp.write(text_tmp1)
    f_temp.close()
    # print(fname(df, type='bill'))
    return f'./files/kopilova.tex'

def pdf(filename, pathname='out'):
    '''
    Функция генерирует pdf-файл из latex-файла

    Вход: 
        filename - str - путь до файла
        pathname - str - папка для файлов
    '''
    print(filename)
    command = ['latexmk', '-pdflatex', f'-outdir={pathname}', filename]

    process = subprocess.Popen(command)

    # process = subprocess.Popen(command)

    process.wait()
    return 0

data = {
    "FirstName": "Анна",
    "MiddleName": "Анатольевна",
    "LastName": "Копылова",
    "SUM": "69000",
    "email": "aa.kopylova@polyketon.ru"
}

# Создание DataFrame
df = pd.Series(data)
print(df["SUM"])
pdf(filename=bill(df), pathname='files/out')