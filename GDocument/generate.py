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

def _apply_tnr_to_run(run, family="Times New Roman", size_pt=None):
    run.font.name = family
    rPr = run._element.rPr
    if rPr is None:
        rPr = run._element.get_or_add_rPr()
    rFonts = rPr.rFonts
    if rFonts is None:
        rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:ascii'), family)
    rFonts.set(qn('w:hAnsi'), family)
    rFonts.set(qn('w:eastAsia'), family)
    rFonts.set(qn('w:cs'), family)


def _apply_tnr_everywhere(doc, family="Times New Roman", size_pt=None):
    # Стиль по умолчанию
    try:
        base = doc.styles["Normal"].font
        base.name = family
        f = doc.styles["Normal"].font.element.rPr.rFonts
        f.set(qn('w:ascii'), family)
        f.set(qn('w:hAnsi'), family)
        f.set(qn('w:eastAsia'), family)
        f.set(qn('w:cs'), family)
    except Exception:
        pass

    # Все ранны в абзацах
    for p in doc.paragraphs:
        for r in p.runs:
            _apply_tnr_to_run(r, family, size_pt)

    # Все ранны в таблицах
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        _apply_tnr_to_run(r, family, size_pt)

def generate_docx(template_path: str, output_path: str, df):
    """
    Генерирует DOCX-файл на основе шаблона, подставляя значения из context.
    
    :param template_path: путь к DOCX-шаблону (например, "template.docx")
    :param output_path: путь к выходному DOCX-файлу (например, "result.docx")
    :param context: словарь с заменами, например {"name": "Иван", "date": "19.09.2025"}
    """
    doc = Document(template_path)
    context = df.to_dict()
    def replace_keys_in_text(text: str) -> str:
        """Заменяет ключи вида {key} на значения из context."""
        for key, value in context.items():
            text = text.replace(f'{{{key}}}', str(value))
        return text
    
    # Проходим по абзацам
    for paragraph in doc.paragraphs:
        if '{' in paragraph.text and '}' in paragraph.text:
            paragraph.text = replace_keys_in_text(paragraph.text)
            # paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Проходим по таблицам
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if '{' in cell.text and '}' in cell.text:
                    cell.text = replace_keys_in_text(cell.text)
    _apply_tnr_everywhere(doc, family="Times New Roman")
    
    doc.save(output_path)



PATTERN = re.compile(r"\{([^{}]+)\}")

# ----------------- ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ -----------------

def _clone_run_format(dst_run: Run, src_run: Run):
    """Скопировать ключевые параметры форматирования из src_run в dst_run."""
    dst_run.bold = src_run.bold
    dst_run.italic = src_run.italic
    dst_run.underline = src_run.underline
    if src_run.font.size is not None:
        dst_run.font.size = src_run.font.size

    if src_run.font.name:
        family = src_run.font.name
        dst_run.font.name = family
        rPr = dst_run._element.get_or_add_rPr()
        rFonts = rPr.get_or_add_rFonts()
        rFonts.set(qn('w:ascii'), family)
        rFonts.set(qn('w:hAnsi'), family)
        rFonts.set(qn('w:eastAsia'), family)
        rFonts.set(qn('w:cs'), family)

def _gather_runs(paragraph) -> List[Tuple[str, Run]]:
    return [(r.text, r) for r in paragraph.runs]

def _split_text_by_placeholders(text: str):
    tokens = []
    last = 0
    for m in PLACEHOLDER_RE.finditer(text):
        if m.start() > last:
            tokens.append(("text", text[last:m.start()]))
        tokens.append(("ph", m.group(1)))
        last = m.end()
    if last < len(text):
        tokens.append(("text", text[last:]))
    return tokens

def _rebuild_paragraph_with_format(
    paragraph,
    run_chunks: List[Tuple[str, Run]],
    context: dict,
    center_keys: bool = False,
    image_key: str = "FILENAME",
    image_width_cm: Optional[float] = None,
    image_height_cm: Optional[float] = None,
    center_image: bool = True,
):
    # Собираем полный текст и карту позиций символов -> исходный run
    full_text = ""
    pos_map: List[Run] = []
    for txt, r in run_chunks:
        full_text += txt
        pos_map.extend([r] * len(txt))

    # Полностью очищаем абзац
    for _ in range(len(paragraph.runs)):
        paragraph.runs[0]._element.getparent().remove(paragraph.runs[0]._element)

    # Флаг: был ли ключ в абзаце (для центрирования)
    has_key_here = False

    idx = 0
    for m in PLACEHOLDER_RE.finditer(full_text):
        # текст до плейсхолдера
        if m.start() > idx:
            seg = full_text[idx:m.start()]
            if seg:
                src_run = pos_map[idx] if idx < len(pos_map) else None
                new_run = paragraph.add_run(seg)
                if src_run is not None:
                    _clone_run_format(new_run, src_run)

        key = m.group(1)
        has_key_here = True

        # ВСТАВКА КАРТИНКИ вместо {FILENAME}
        if key == image_key:
            path = context.get(key)
            # создаём "пустой" run, чтобы вставить в него картинку
            src_run_for_ph: Optional[Run] = pos_map[m.start()] if m.start() < len(pos_map) else None
            pic_run = paragraph.add_run()
            if src_run_for_ph is not None:
                _clone_run_format(pic_run, src_run_for_ph)
            # размеры
            kw = {}
            if image_width_cm is not None:
                kw["width"] = Cm(image_width_cm)
            if image_height_cm is not None:
                kw["height"] = Cm(image_height_cm)
            # вставка (если путь не задан — оставляем плейсхолдер как текст)
            if path:
                pic_run.add_picture(path, **kw)
                # выравнивание абзаца под картинку при необходимости
                if center_image:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                # нет пути — подставим исходный текст плейсхолдера
                fallback_run = paragraph.add_run("{"+key+"}")
                if src_run_for_ph is not None:
                    _clone_run_format(fallback_run, src_run_for_ph)
                else:
                    # обычная текстовая подстановка
                    value = str(context.get(key, "{"+key+"}"))
                    src_run_for_ph: Optional[Run] = pos_map[m.start()] if m.start() < len(pos_map) else None
                    ph_run = paragraph.add_run(value)
                    if src_run_for_ph is not None:
                        _clone_run_format(ph_run, src_run_for_ph)

                idx = m.end()

    # хвост текста после последнего плейсхолдера
    if idx < len(full_text):
        seg = full_text[idx:]
        src_run = pos_map[idx] if idx < len(pos_map) else None
        new_run = paragraph.add_run(seg)
        if src_run is not None:
            _clone_run_format(new_run, src_run)

    # Центрировать абзац, если в нём были ключи и это запрошено
    if has_key_here and center_keys:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

def _process_paragraph(
    paragraph,
    context: dict,
    center_keys: bool = False,
    image_key: str = "FILENAME",
    image_width_cm: Optional[float] = None,
    image_height_cm: Optional[float] = None,
    center_image: bool = True,
):
    if "{" not in paragraph.text or "}" not in paragraph.text:
        return
    chunks = _gather_runs(paragraph)
    if len(chunks) == 1:
        # быстрый путь
        txt, r = chunks[0]
        tokens = _split_text_by_placeholders(txt)
        paragraph.clear()
        has_key_here = False
        for kind, payload in tokens:
            if kind == "text":
                run = paragraph.add_run(payload)
                _clone_run_format(run, r)
            else:  # ph
                has_key_here = True
                key = payload
                if key == image_key:
                    path = context.get(key)
                    pic_run = paragraph.add_run()
                    _clone_run_format(pic_run, r)
                    if path:
                        kw = {}
                        if image_width_cm is not None:
                            kw["width"] = Cm(image_width_cm)
                        if image_height_cm is not None:
                            kw["height"] = Cm(image_height_cm)
                        pic_run.add_picture(path, **kw)
                        if center_image:
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    else:
                        fb = paragraph.add_run("{"+key+"}")
                        _clone_run_format(fb, r)
                else:
                    val = str(context.get(key, "{"+key+"}"))
                    rr = paragraph.add_run(val)
                    _clone_run_format(rr, r)
        if has_key_here and center_keys:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        return

    # сложный случай: плейсхолдеры пересекают несколько run'ов
    _rebuild_paragraph_with_format(
        paragraph, chunks, context,
        center_keys=center_keys,
        image_key=image_key,
        image_width_cm=image_width_cm,
        image_height_cm=image_height_cm,
        center_image=center_image
    )

def _process_cell(cell, **kwargs):
    for p in cell.paragraphs:
        _process_paragraph(p, **kwargs)

# ----------------- ОСНОВНАЯ ФУНКЦИЯ -----------------

def generate_docx_preserve_format(
    template_path: str,
    output_path: str,
    context: pd.Series,
    *,
    center_keys: bool = False,
    image_key: str = "FILENAME",
    image_width_cm: Optional[float] = None,
    image_height_cm: Optional[float] = None,
    center_image: bool = True,
):
    """
    Подстановка значений из pd.Series с сохранением форматирования шаблона.
    Специальный ключ {FILENAME} (или другой через image_key) вставляет картинку.
    - center_keys=True центрирует абзацы, содержащие ключи.
    - image_width_cm/height_cm — опциональные размеры изображения.
    - center_image=True — центрировать абзац с картинкой.
    """
    doc = Document(template_path)
    ctx = context.to_dict()
    for p in doc.paragraphs:
        _process_paragraph(
            p, ctx,
            center_keys=center_keys,
            image_key=image_key,
            image_width_cm=image_width_cm,
            image_height_cm=image_height_cm,
            center_image=center_image
        )

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                _process_cell(
                    cell,
                    context=ctx,
                    center_keys=center_keys,
                    image_key=image_key,
                    image_width_cm=image_width_cm,
                    image_height_cm=image_height_cm,
                    center_image=center_image
                )

    doc.save(output_path)

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

    fileObj = codecs.open( "templates/bill.tex", "r", "utf_8" )
    text = fileObj.read()

    text_tmp = text.replace('NAMEORG', NameOrg)
    text_tmp = text_tmp.replace('INN', PayeeINN)
    text_tmp = text_tmp.replace('NUMBERCHET', PersonalAcc)
    text_tmp = text_tmp.replace('NAMEBANK', BankName)
    text_tmp = text_tmp.replace('BICBANK', BIC)
    text_tmp = text_tmp.replace('KORRCHET', CorrespAcc)
    text_tmp = text_tmp.replace('KPP', KPP)

    # text_tmp = text_tmp.replace('DOCNUM', df['DOCNUM'])

    if df["MIDDLE_NAME"]!=df["MIDDLE_NAME"]:
        text_tmp = text_tmp.replace('PLATE', f'{df["LAST_NAME"]} {df["FIRST_NAME"]}')
        filename = f'{df['LAST_NAME']}_{df["FIRST_NAME"]}_{df["SUM"]}'
    else:
        text_tmp = text_tmp.replace('PLATE', f'{df["LAST_NAME"]} {df["FIRST_NAME"]} {df["MIDDLE_NAME"]}')
        filename = f'{df['LAST_NAME']}_{df["FIRST_NAME"][0]}_{df["MIDDLE_NAME"][0]}_{df["SUM"]}'

    text_tmp1 = text_tmp.replace('SUM', df['SUMM'])
    text_tmp1 = text_tmp1.replace('FILENAME.png', qr_code(df))

    f_temp = codecs.open(f'files/{fname(df, type='bill')}.tex', 'w', "utf_8")
    f_temp.write(text_tmp1)
    f_temp.close()
    # print(fname(df, type='bill'))
    return f'./files/{fname(df, type='bill')}.tex'

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


def act(df):
    '''
    Функция генерирует tex-файл для акта по шаблону
    Шаблон счета расположен: /templates/act.tex
    Путь к файлу: /files
    Название файла генерируется при помощи функции fname с параметром type='bill'
    Вход:
        df - pd.Series - информация об участнике
    Выход:
        Путь к сохраненному файлу 

    '''
    fileObj = codecs.open("templates/act.tex", "r", "utf_8" )
    text = fileObj.read()
    # text_tmp = text.replace('DOCNUM', df['DOCNUM'])
    
    text_tmp = text.replace('FIRST_NAME', df['FIRST_NAME'])


    if df['SUMM']=='nan':
        text_tmp = text_tmp.replace('SUMM', '\\underline{\parbox{1 cm}{\hphantom}}')
        text_tmp = text_tmp.replace('SNAME', '\\underline{\parbox{5 cm}{\hphantom}}')
    else:
        # print(float(df['SUMM'])/1000)
        text_tmp = text_tmp.replace('SUMM', str(int(df['SUMM'])//1000))
        text_tmp = text_tmp.replace('SNAME', num2words(int(df['SUMM'])//1000, lang='ru'))
        

    if df['SEX']=='m':
        text_tmp = text_tmp.replace('SEX', 'ый')
    elif df['SEX']=='w':
        text_tmp = text_tmp.replace('SEX', 'ая')
    else:
        text_tmp = text_tmp.replace('SEX', 'ый(ая)')

    # if (df['MIDDLE_NAME']=='nan') or ((df['MIDDLE_NAME'])!= (df['MIDDLE_NAME'])):
    #     text_tmp = text_tmp.replace('FNAME.', df['LAST_NAME'])
    #     text_tmp = text_tmp.replace('MNAME.', '')
    #     text_tmp = text_tmp.replace('LAST_NAME.', df['FIRST_NAME'][0])
    #     text_tmp = text_tmp.replace('MIDDLE_NAME', '')
        
    # else:
    #     text_tmp = text_tmp.replace('FNAME', df['FIRST_NAME'][0])
    #     text_tmp = text_tmp.replace('MIDDLE_NAME', df['MIDDLE_NAME'])
    #     text_tmp = text_tmp.replace('MNAME', df['MIDDLE_NAME'][0])
    #     text_tmp = text_tmp.replace('LAST_NAME.', df['LAST_NAME'])
    text_tmp = text_tmp.replace('FNAME', df['FIRST_NAME'][0])
    text_tmp = text_tmp.replace('LAST_NAME', df['LAST_NAME'])
    if (df['MIDDLE_NAME']=='nan') or ((df['MIDDLE_NAME'])!=(df['MIDDLE_NAME'])):
        text_tmp = text_tmp.replace('MNAME.', '')
        text_tmp = text_tmp.replace('MIDDLE_NAME', '')
    else:
        text_tmp = text_tmp.replace('MIDDLE_NAME', df['MIDDLE_NAME'])
        text_tmp = text_tmp.replace('MNAME', df['MIDDLE_NAME'][0])
        

    f_temp = codecs.open(f'./files/{fname(df, type='act')}.tex', 'w', "utf_8")

    text_tmp = text_tmp.replace('LAST_NAME', df['LAST_NAME'])
    f_temp.write(text_tmp)
    f_temp.close()
    return f'./files/{fname(df, type='act')}.tex'

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
    df1['SUMM_NAME'] = num2words(int(df1['SUMM']), lang='ru')

    df1['F_NAME'] = df1['FIRST_NAME'].str[0] + '.'
    df1['M_NAME'] = df1['MIDDLE_NAME'].str[0] + '.'
    df1["FILENAE"] = df1.apply(lambda row: fname(row, 'qr')+ ".png", axis =1)
    # df1["FILENAE"] = df1.apply(lambda row: num2words(row, 'qr')+ ".png", axis =1)

    # print(df1)
    for person_ID in range(len(df1)):
        print()
        
        qr_code(df1.iloc[person_ID])
        # generate_docx(f'{config.TEMP_FOLDER_NAME}/act.docx', f'{path}/out/{fname(df1.iloc[person_ID], 'act')}'+ '.docx', df1.iloc[person_ID])
        print(df1.iloc[person_ID])
        # generate_docx(f'{config.TEMP_FOLDER_NAME}/bill.docx', f'{path}/out/{fname(df1.iloc[person_ID], 'bill')}'+ '.docx', df1.iloc[person_ID])

        generate_docx_from_template(f'{config.TEMP_FOLDER_NAME}/act.docx',f'{path}/out/{fname(df1.iloc[person_ID], 'act')}'+ '.docx', df1.iloc[person_ID])

    return 'qr code создан'
        

    
    
    # print(bill(df))
    # pdf(filename=bill(df), pathname=f'{path}/out')
    # pdf(filename=act(df), pathname=f'{path}/out')



def gen_all_summ(df, out_dir = 'files'):
    '''
    Функция запускает процесс генерации договора и счета для всех вариантов сумм
    Вход: 
        df - pd.Series - информация об участнике
    '''
    summ = ['80000', '60000']
    if not os.path.isdir(out_dir+'/out'):
        os.makedirs(out_dir+'/out',exist_ok=True)

    if not os.path.isdir(out_dir+'qr_code'):
        os.makedirs(out_dir+'/qr_code',exist_ok=True)
    for i in summ:
        df['SUMM'] = i
        # print(df)
    # print(bill(df))
        pdf(filename=bill(df), pathname=out_dir+'/out')
        pdf(filename=act(df), pathname=out_dir+'/out') 