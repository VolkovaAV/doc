# from . import NameOrg, PersonalAcc, BankName, BIC, CorrespAcc, KPP, PayeeINN
from docx import Document
from docx.oxml.ns import qn

import subprocess # для запуска процессов 
from transliterate import translit # для корректной записи файлов

import qrcode # для генерации кода
import codecs # для обработки рускоязычных файлов

from num2words import num2words 

import copy

import os

import config
import pandas as pd
import numpy as np


import re


from .create import create_excel_with_columns


from typing import List, Tuple, Optional

from docx.text.run import Run

from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH


from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import os
from typing import Dict, Any, Optional

NameOrg = 'МЦФПИН'
PersonalAcc = '40703810942000000672'
BankName= 'ВОЛГО-ВЯТСКИЙ БАНК ПАО СБЕРБАНК'
BIC = '042202603'
CorrespAcc='30101810900000000603'
KPP='526001001'
PayeeINN='5260054053'

from_mail = "pay.incas@mail.ru"                         # Почта отправителя
from_passwd = "8sRt0H0xDvqTmCxhNGx8"                      # пароль от почты отправителя
server_adr = "smtp.mail.ru"                               # адрес почтового сервера

to_mail_test = 'a.volkova@ipfran.ru'                           # адрес получателя

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

def fname(df, type):
    '''
    Генерирует название файла. ФИО участника транслитеруется на латиницу, ь пропускается
    Вход: 
        df - pd.Series - информация об участнике
        type - str - тип файла (н-р, contract для договора и bill для счета)
    Выход:
        result - str - название файла
    '''
    # custom_replacements = {('ь', '')}
    if df["MIDDLE_NAME"]!=df["MIDDLE_NAME"]:
        name = translit(df['LAST_NAME']+'_'+df['FIRST_NAME'], language_code='ru', reversed=True, strict=True)
    else:
        name = translit(df['LAST_NAME'] + '_' + df['FIRST_NAME'][0] + '_' + df['MIDDLE_NAME'][0], language_code='ru', reversed=True, strict=True)
        # name = translit('фф', language_code='ru', strict=True, )
    name = name.replace("'", "")
    result = name  + '_' + type +'_'+ str(df['SUMM']) 

    return result

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

    if df["MIDDLE_NAME"]!=df["MIDDLE_NAME"]:
        PersonInfo = '(участник '+df['LAST_NAME']+' '+df['FIRST_NAME'][0]+'.)'
    else:
        PersonInfo = '(участник '+df['LAST_NAME']+' '+df['FIRST_NAME'][0]+'. '+df['MIDDLE_NAME'][0]+'.)'

    data=f'ST00012|'\
        f'Name={NameOrg}|'\
        f'PersonalAcc={PersonalAcc}|'\
        f'BankName={BankName}|'\
        f'BIC={BIC}|'\
        f'CorrespAcc={CorrespAcc}|'\
        f'KPP={KPP}|'\
        f'PayeeINN={PayeeINN}|'\
        f'Purpose= {config.PAY_PURPOSE} {PersonInfo}|'\
        f'SUM={int(df['SUMM'])*100}'
    
    img = qrcode.make(data)
    path_name = f'./files/qr_code/{fname(df, type='qr')}.png'
    img.save(path_name)
    return path_name

def _series_to_dict(ctx: pd.Series | Dict[str, str]) -> Dict[str, str]:
    if isinstance(ctx, pd.Series):
        # Преобразуем к строкам, чтобы избежать "nan"
        return {str(k): ("" if pd.isna(v) else str(v)) for k, v in ctx.items()}
    return {str(k): ("" if v is None else str(v)) for k, v in ctx.items()}

def generate_docx_advanced(
    template_path: str,
    output_path: str,
    df: pd.Series,
    image_mapping: Optional[Dict[str, str]] = None,
    default_image_width: int = 60
) -> None:
    """
    Усовершенствованная версия с явным указанием mapping изображений.
    
    Args:
        image_mapping: Словарь {ключ_в_шаблоне: путь_к_изображению}
    """
    context = _series_to_dict(df)
    image_mapping = {"FILENAME": f"files/qr_code/{fname(df, 'qr')}.png"}
    try:
        doc = DocxTemplate(template_path)
        
        # Обрабатываем изображения
        if image_mapping:
            for key, image_path in image_mapping.items():
                if os.path.exists(image_path):
                    image = InlineImage(doc, image_path, width=Mm(default_image_width))
                    context[key] = image
                else:
                    print(f"Предупреждение: изображение не найдено - {image_path}")
                    context[key] = "[Изображение не найдено]"
        
        doc.render(context)
        
        # Создаем папку для output если не существует
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        doc.save(output_path)
        
        return f"Документ создан: {output_path}"
        
    except Exception as e:
        print(f"Ошибка: {e}")
        raise

def email(df, first=True):
    '''
    Функция генерирует текст для электронного письма с использованием html-разметки

    Вход:
        df - pd.Series - информация об участнике
        first - bool - проверка на статус (если true используется шаблон: /templates/email_first.html, если false: /templates/email_resent.html)
    Выход:
        text - str - текст письма
    '''
    if first:
        suffix='first'
    else:
        suffix='resent'
    with open(f"templates/email_{suffix}.html", encoding='utf-8') as f:
        text_template = f.read()

    text = text_template.replace('FIRST_NAME', df['FIRST_NAME'])
    if (df['MIDDLE_NAME'] != df['MIDDLE_NAME']) or (len(df['MIDDLE_NAME']) < 2):
        text = text.replace('MIDDLE_NAME', df['LAST_NAME'])
    else:
        text = text.replace('MIDDLE_NAME', df['MIDDLE_NAME'])
    
    if df['SEX']=='m':
        text = text.replace('SEX', 'ый')
    elif df['SEX']=='w':
        text = text.replace('SEX', 'ая')
    else:
        text = text.replace('SEX', 'ый(ая)')
    
    return text

def checking_exel(name_table):
    if os.path.exists(name_table):
        print(f"⚠️ Файл '{name_table}' уже существует. Создание отменено.")
        return True
    else:
        return False

def gen_all(path):
    '''
    Функция запускает процесс генерации акта и счета
    Вход: 
        df - pd.Series - информация об участнике
    '''
    if not os.path.isdir(f'{path}/out'):
        os.makedirs(f'{path}/out')

    if not os.path.isdir(f'{path}/qr_code'):
        os.makedirs(f'{path}/qr_code')

    xl = pd.read_excel(config.TB_NAME, dtype='str')
    df1 = pd.DataFrame(xl)

    df1 =df1.rename(columns={'Фамилия': 'LAST_NAME', 'Имя': 'FIRST_NAME', 'Отчество': 'MIDDLE_NAME', 'Сумма': 'SUMM'})
    
    df1['SEX'] = np.where(df1['MIDDLE_NAME'].str.endswith('на'), 'ая',np.where(df1['MIDDLE_NAME'].str.endswith('ич'), 'ый','ый(ая)')
)
    df1['SUMM_NAME'] = df1['SUMM'].apply(lambda x: num2words(int(x), lang='ru'))

    df1['F_NAME'] = df1['FIRST_NAME'].str[0] + '.'
    df1['M_NAME'] = df1['MIDDLE_NAME'].str[0] + '.'


    # print(df1)
    for person_ID in range(len(df1)):
        res = ''
        qr_code(df1.iloc[person_ID])+ '\n'
        generate_docx_advanced(f'{config.TEMP_FOLDER_NAME}/bill.docx', f'{path}/out/{fname(df1.iloc[person_ID], 'bill')}.docx', df1.iloc[person_ID])
        generate_docx_advanced(f'{config.TEMP_FOLDER_NAME}/act.docx', f'{path}/out/{fname(df1.iloc[person_ID], 'act')}.docx', df1.iloc[person_ID]) + '\n'

    res += 'Генерация завершена!' 

    return res
        

    
    
    # print(bill(df))
    # pdf(filename=bill(df), pathname=f'{path}/out')
    # pdf(filename=act(df), pathname=f'{path}/out')
