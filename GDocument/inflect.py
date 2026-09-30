"""
Склонение названия мероприятия по падежам (pymorphy3).

Название вводится в именительном падеже, например «Научная школа».
Склоняется начало фразы до главного существительного включительно:
прилагательные согласуются с ним, а всё, что стоит после
(«молодых учёных», «по физике», «"Актуальные проблемы..."»), не меняется.

    >>> inflect_phrase('Научная школа', 'loct')
    'Научной школе'
    >>> inflect_phrase('Научная школа', 'gent')
    'Научной школы'
"""
import re
from functools import lru_cache

import pymorphy3
import pymorphy3_dicts_ru

PREP_CASE = 'loct'   # предложный: (в) чём?
GEN_CASE = 'gent'    # родительный: чего?

_INFLECTABLE = {'NOUN', 'ADJF', 'PRTF'}
# слово = [открывающая пунктуация][буквы/цифры/дефисы][закрывающая пунктуация]
_TOKEN = re.compile(r'^(?P<pre>[^\w]*)(?P<word>[\w-]+?)(?P<post>[^\w]*)$')


@lru_cache(maxsize=1)
def _morph():
    # путь к словарям указываем явно, чтобы работало и внутри exe (PyInstaller)
    return pymorphy3.MorphAnalyzer(path=pymorphy3_dicts_ru.get_path())

def _nominative_parse(word):
    """Разбор слова как существительного/прилагательного в им. падеже, иначе None."""
    for p in _morph().parse(word):
        if p.tag.POS in _INFLECTABLE and 'nomn' in p.tag.case:
            return p
    return None

def _keep_style(original, inflected):
    """Переносит регистр и написание е/ё исходного слова на склоненное."""
    if 'ё' not in original.lower():
        inflected = inflected.replace('ё', 'е')
    if original.isupper() and len(original) > 1:
        return inflected.upper()
    if original[:1].isupper():
        # сохраняем заглавные в частях через дефис: «Школа-Конференция»
        parts_o, parts_i = original.split('-'), inflected.split('-')
        if len(parts_o) == len(parts_i):
            return '-'.join(i.capitalize() if o[:1].isupper() else i for o, i in zip(parts_o, parts_i))
        return inflected.capitalize()
    return inflected

def inflect_phrase(phrase: str, case: str) -> str:
    """
    Склоняет название в именительном падеже в падеж case ('loct', 'gent', ...).
    Если слово распознать не удалось, оно остается без изменений.
    """
    tokens = phrase.split()
    adjectives = []   # (индекс, разбор) прилагательных до главного слова
    head = None       # (индекс, разбор) главного существительного

    for i, token in enumerate(tokens):
        m = _TOKEN.match(token)
        # кавычки открывают собственное название — его не склоняем
        if not m or m.group('pre') or any(c in token for c in '«"\''):
            break
        word = m.group('word')
        if re.fullmatch(r'[IVXLCDM]+|\d[\w-]*', word):  # «V», «XV», «2025»
            continue
        p = _nominative_parse(word)
        if p is None:
            break
        if p.tag.POS == 'NOUN':
            head = (i, p)
            break
        adjectives.append((i, p))
        if m.group('post'):  # запятая и т.п. после прилагательного
            break

    def put(i, parse, grammemes):
        m = _TOKEN.match(tokens[i])
        form = parse.inflect(grammemes)
        if form is not None:
            tokens[i] = m.group('pre') + _keep_style(m.group('word'), form.word) + m.group('post')

    if head is not None:
        _, noun = head
        number = noun.tag.number or 'sing'
        agree = {case, number}
        if number == 'sing' and noun.tag.gender:
            agree.add(noun.tag.gender)
        for i, p in adjectives:
            put(i, p, agree)
        put(head[0], noun, {case})
    else:
        for i, p in adjectives:
            put(i, p, {case})

    return ' '.join(tokens)

def event_cases(event_info: str) -> dict:
    """Название мероприятия в предложном и родительном падежах."""
    event_info_prep = inflect_phrase(event_info, PREP_CASE)  # (в) Научной школе
    event_info_gen = inflect_phrase(event_info, GEN_CASE)    # (для) Научной школы
    return {'EVENT_INFO_PREP': event_info_prep, 'EVENT_INFO_GEN': event_info_gen}
