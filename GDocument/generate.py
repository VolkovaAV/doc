from transliterate import translit # для корректной записи файлов

from docx2pdf import convert

import qrcode # для генерации кода

from num2words import num2words

import html
import os
from string import Template
from typing import Dict, Optional

import config
from .json_work import *

import pandas as pd
import numpy as np

from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm


OUT_DIR = os.path.join(config.FILES_FOLDER_NAME, 'out')
PDF_DIR = os.path.join(config.FILES_FOLDER_NAME, 'pdf')
QR_DIR = os.path.join(config.FILES_FOLDER_NAME, 'qr_code')


def has_middle_name(df):
    '''Есть ли у участника отчество (пустые ячейки Excel читаются как NaN)'''
    return isinstance(df['MIDDLE_NAME'], str) and df['MIDDLE_NAME'].strip() != ''

def load_participants():
    '''
    Читает таблицу участников и добавляет вычисляемые столбцы.
    Выход:
        df - pd.DataFrame - по строке на участника
    '''
    if not os.path.isfile(config.TB_NAME):
        raise FileNotFoundError(f"Не найден файл участников '{config.TB_NAME}'. Нажмите «Создать шаблоны».")

    df = pd.read_excel(config.TB_NAME, dtype='str')
    df = df.rename(columns={'Фамилия': 'LAST_NAME', 'Имя': 'FIRST_NAME', 'Отчество': 'MIDDLE_NAME', 'Сумма': 'SUMM'})

    # полностью пустые строки (частый случай в Excel) пропускаем
    df = df.dropna(how='all').reset_index(drop=True)

    for col in ('LAST_NAME', 'FIRST_NAME', 'email', 'SUMM'):
        empty = df[col].isna() | (df[col].str.strip() == '')
        if empty.any():
            rows = ', '.join(str(i + 2) for i in df.index[empty])  # +2: заголовок и нумерация Excel с 1
            raise ValueError(f"В '{config.TB_NAME}' не заполнен столбец '{col}' (строки: {rows})")
        df[col] = df[col].str.strip()

    bad_summ = ~df['SUMM'].str.fullmatch(r'\d+')
    if bad_summ.any():
        rows = ', '.join(str(i + 2) for i in df.index[bad_summ])
        raise ValueError(f"В '{config.TB_NAME}' сумма должна быть целым числом рублей (строки: {rows})")

    middle = df['MIDDLE_NAME'].fillna('').str.strip()
    df['MIDDLE_NAME'] = middle.replace('', np.nan)
    df['SEX'] = np.where(middle.str.endswith('на'), 'ая', np.where(middle.str.endswith('ич'), 'ый', 'ый(ая)'))

    df['F_NAME'] = df['FIRST_NAME'].str[0] + '.'
    df['M_NAME'] = np.where(middle != '', middle.str[0] + '.', '')
    return df

def fname(df, type):
    '''
    Генерирует название файла. ФИО участника транслитеруется на латиницу, ь пропускается
    Вход:
        df - pd.Series - информация об участнике
        type - str - тип файла (н-р, contract для договора и bill для счета)
    Выход:
        result - str - название файла
    '''
    if not has_middle_name(df):
        name = translit(df['LAST_NAME']+'_'+df['FIRST_NAME'], language_code='ru', reversed=True, strict=True)
    else:
        name = translit(df['LAST_NAME'] + '_' + df['FIRST_NAME'][0] + '_' + df['MIDDLE_NAME'][0], language_code='ru', reversed=True, strict=True)
    name = name.replace("'", "")
    result = name  + '_' + type +'_'+ str(df['SUMM'])

    return result

def qr_code(df, params):
    '''
    Функция сохраняет в папку проекта картинку с qr кодом для оплаты.
    Картинка сохраняется в папку QR_DIR
    Название файла генерируется при помощи функции fname() с параметром type='qr'

    Вход:
        df - pd.Series - информация об участнике
    Выход:
        path_name - название сохраненного файла
    '''
    PAY_PURPOSE = f"Рег. взнос за участие в {params['EVENT_NAME']}, {params['DATE_INFO']}, {params['PLACE_INFO']}"

    if not has_middle_name(df):
        PersonInfo = '(участник '+df['LAST_NAME']+' '+df['FIRST_NAME'][0]+'.)'
    else:
        PersonInfo = '(участник '+df['LAST_NAME']+' '+df['FIRST_NAME'][0]+'. '+df['MIDDLE_NAME'][0]+'.)'

    data=f'ST00012|'\
        f'Name={config.NameOrg}|'\
        f'PersonalAcc={config.PersonalAcc}|'\
        f'BankName={config.BankName}|'\
        f'BIC={config.BIC}|'\
        f'CorrespAcc={config.CorrespAcc}|'\
        f'KPP={config.KPP}|'\
        f'PayeeINN={config.PayeeINN}|'\
        f'Purpose= {PAY_PURPOSE} {PersonInfo}|'\
        f"SUM={int(df['SUMM'])*100}"

    img = qrcode.make(data)
    path_name = os.path.join(QR_DIR, fname(df, type='qr') + '.png')
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
) -> str:
    """
    Заполняет docx-шаблон данными участника.

    Args:
        image_mapping: Словарь {ключ_в_шаблоне: путь_к_изображению}.
                       По умолчанию в {{ FILENAME }} вставляется QR-код участника.
    """
    context = _series_to_dict(df)
    if image_mapping is None:
        image_mapping = {"FILENAME": os.path.join(QR_DIR, fname(df, 'qr') + '.png')}

    doc = DocxTemplate(template_path)

    # Обрабатываем изображения
    for key, image_path in image_mapping.items():
        if not os.path.exists(image_path):
            raise FileNotFoundError(f"Изображение не найдено: {image_path}")
        context[key] = InlineImage(doc, image_path, width=Mm(default_image_width))

    doc.render(context)

    # Создаем папку для output если не существует
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)

    return f"Документ создан: {output_path}"

def pdf(df, type):
    os.makedirs(PDF_DIR, exist_ok=True)

    convert(os.path.join(OUT_DIR, fname(df, type) + ".docx"), os.path.join(PDF_DIR, fname(df, type) + ".pdf"))
    return f"{df['LAST_NAME']} {df['FIRST_NAME']}: генерация завершена"

def email(df):
    '''
    Функция генерирует текст для электронного письма с использованием html-разметки

    Вход:
        df - pd.Series - информация об участнике
    Выход:
        text - str - текст письма
    '''

    with open(os.path.join(config.TEMP_FOLDER_NAME, "email.html"), encoding='utf-8') as f:
        text_template = f.read()

    if '${FIRST_NAME}' not in text_template:
        raise ValueError("Шаблон письма устарел: нажмите «Создать шаблоны», чтобы пересоздать templates/email.html")

    if not has_middle_name(df) or len(df['MIDDLE_NAME']) < 2:
        middle = df['LAST_NAME']
    else:
        middle = df['MIDDLE_NAME']

    return Template(text_template).safe_substitute(
        SEX=df['SEX'],
        FIRST_NAME=html.escape(df['FIRST_NAME']),
        MIDDLE_NAME=html.escape(middle),
    )

def generate_one_person(df, params):
    qr_code(df, params)
    generate_docx_advanced(os.path.join(config.TEMP_FOLDER_NAME, 'bill.docx'),
                           os.path.join(OUT_DIR, fname(df, 'bill') + '.docx'), df)
    return pdf(df, 'bill')

def gen_all():
    '''
    Функция запускает генерацию счетов для всех участников из таблицы
    '''
    params = load_config()
    require_event_params(params)

    template = os.path.join(config.TEMP_FOLDER_NAME, 'bill.docx')
    if not os.path.isfile(template):
        raise FileNotFoundError(f"Не найден шаблон '{template}'. Нажмите «Создать шаблоны».")

    os.makedirs(OUT_DIR, exist_ok=True)
    os.makedirs(QR_DIR, exist_ok=True)

    df1 = load_participants()
    df1['SUMM_NAME'] = df1['SUMM'].apply(lambda x: num2words(int(x), lang='ru'))

    lines = [generate_one_person(df1.iloc[person_ID], params) for person_ID in range(len(df1))]
    lines.append(f'Генерация завершена! Участников: {len(df1)}')

    return '\n'.join(lines)
