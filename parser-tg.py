import re
import argparse
import pandas as pd
from datetime import datetime, timedelta


def parse_date(raw: str) -> str:
    """
    Преобразует различные форматы дат в единый формат 'DD.MM.YYYY HH:MM'.
    Поддерживает:
        - полные даты 'DD.MM.YYYY HH:MM'
        - 'X минут(а/ы) назад'
        - 'час назад', 'два часа назад', 'три часа назад'
        - 'X часов назад'
        - 'сегодня в HH:MM'
        - 'вчера в HH:MM'
        - 'DD мес. в HH:MM'
    Если не удалось распознать, возвращает исходную строку.
    """
    if not raw:
        return ''
    now = datetime.now()
    raw_l = raw.lower().strip()

    # Полная дата 'DD.MM.YYYY HH:MM'
    m = re.match(r"(\d{2})\.(\d{2})\.(\d{4})\s+(\d{1,2}):(\d{2})", raw_l)
    if m:
        day, month, year, hour, minute = map(int, m.groups())
        dt = datetime(year, month, day, hour, minute)
        return dt.strftime('%d.%m.%Y %H:%M')

    # X минут назад
    m = re.match(r"(\d+)\s+минут", raw_l)
    if m:
        dt = now - timedelta(minutes=int(m.group(1)))
        return dt.strftime('%d.%m.%Y %H:%M')

    # текстовые варианты часов: 'час назад', 'два часа назад', 'три часа назад'
    textual_map = {'час назад': 1, 'два часа назад': 2, 'три часа назад': 3}
    if raw_l in textual_map:
        dt = now - timedelta(hours=textual_map[raw_l])
        return dt.strftime('%d.%m.%Y %H:%M')

    # цифровые часы назад 'X часов назад'
    m = re.match(r"(\d+)\s+час", raw_l)
    if m:
        dt = now - timedelta(hours=int(m.group(1)))
        return dt.strftime('%d.%m.%Y %H:%M')

    # сегодня в HH:MM
    m = re.match(r"сегодня в (\d{1,2}):(\d{2})", raw_l)
    if m:
        hour, minute = int(m.group(1)), int(m.group(2))
        dt = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
        return dt.strftime('%d.%m.%Y %H:%M')

    # вчера в HH:MM
    m = re.match(r"вчера в (\d{1,2}):(\d{2})", raw_l)
    if m:
        hour, minute = int(m.group(1)), int(m.group(2))
        dt = (now - timedelta(days=1)).replace(hour=hour, minute=minute, second=0, microsecond=0)
        return dt.strftime('%d.%m.%Y %H:%M')

    # DD мес. в HH:MM
    mon_map = {
        'янв': 1, 'фев': 2, 'мар': 3, 'апр': 4, 'май': 5, 'июн': 6,
        'июл': 7, 'авг': 8, 'сен': 9, 'окт': 10, 'ноя': 11, 'дек': 12
    }
    m = re.match(r"(\d{1,2})\s+(янв|фев|мар|апр|май|июн|июл|авг|сен|окт|ноя|дек) в (\d{1,2}):(\d{2})", raw_l)
    if m:
        day = int(m.group(1))
        month = mon_map[m.group(2)]
        hour, minute = int(m.group(3)), int(m.group(4))
        year = now.year
        dt = datetime(year, month, day, hour, minute)
        return dt.strftime('%d.%m.%Y %H:%M')

    # fallback
    return raw


def parse_votes(md_file: str) -> pd.DataFrame:
    """
    Читает файл с блоками:
      Имя, [дата]
      комментарий
    и извлекает имя, номер участника и дату.
    """
    header_re = re.compile(r"^(.+), \[(\d{2}\.\d{2}\.\d{4} \d{1,2}:\d{2})\]$")
    number_re = re.compile(r"([1-9][0-9]?)")

    records = []
    with open(md_file, 'r', encoding='utf-8') as f:
        lines = [line.strip() for line in f]

    i = 0
    while i < len(lines):
        m = header_re.match(lines[i])
        if m and i + 1 < len(lines):
            username = m.group(1)
            raw_date = m.group(2)
            comment = lines[i + 1]
            num_match = number_re.search(comment)
            if num_match:
                participant = int(num_match.group(1))
                records.append({'Имя пользователя': username,
                                'Номер участника': participant,
                                'Дата голосования': raw_date})
            i += 2
        else:
            i += 1

    df = pd.DataFrame(records)
    df['Дата и время (Excel)'] = df['Дата голосования'].apply(parse_date)
    df['Дата'] = df['Дата и время (Excel)'].apply(lambda x: x.split(' ')[0] if ' ' in x else x)
    return df


def summarize_all(df: pd.DataFrame) -> pd.DataFrame:
    return (df['Номер участника']
            .value_counts()
            .sort_index()
            .rename_axis('Номер участника')
            .reset_index(name='Количество голосов'))


def detail_by_user(df: pd.DataFrame) -> pd.DataFrame:
    return df[['Имя пользователя', 'Номер участника', 'Дата голосования', 'Дата и время (Excel)', 'Дата']]


def summarize_unique(df: pd.DataFrame) -> pd.DataFrame:
    return (df.drop_duplicates(subset=['Имя пользователя'], keep='first')['Номер участника']
            .value_counts()
            .sort_index()
            .rename_axis('Номер участника')
            .reset_index(name='Уникальные голоса'))


def main():
    parser = argparse.ArgumentParser(description='Подсчет голосов из текстового файла')
    parser.add_argument('input_file', help='Путь к файлу с голосами')
    parser.add_argument('--excel', action='store_true', help='Сохранить результаты в Excel (results-tg.xlsx)')
    args = parser.parse_args()

    print('*Парсинг голосов начат!* 😊')
    df = parse_votes(args.input_file)

    print('\n*Таблица №1 — Сводная по всем голосам:*')
    print(summarize_all(df).to_string(index=False))

    print('\n*Таблица №2 — Детализация по пользователям:*')
    print(detail_by_user(df).to_string(index=False))

    print('\n*Таблица №3 — Сводная по уникальным голосам:*')
    print(summarize_unique(df).to_string(index=False))

    if args.excel:
        with pd.ExcelWriter('results-tg.xlsx') as writer:
            summarize_all(df).to_excel(writer, sheet_name='Все голоса', index=False)
            detail_by_user(df).to_excel(writer, sheet_name='Детализация', index=False)
            summarize_unique(df).to_excel(writer, sheet_name='Уникальные голоса', index=False)
        print('\n**Результаты сохранены в results-tg.xlsx** 🎉')


if __name__ == '__main__':
    main()
