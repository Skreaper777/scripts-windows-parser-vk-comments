import re
import argparse
import pandas as pd
from datetime import datetime, timedelta


def parse_date(raw: str) -> str:
    """
    Преобразует различные форматы дат ВКонтакте в единый формат 'DD.MM.YYYY HH:MM'.
    Поддерживает:
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
    # fallback: возвращаем исходную строку
    return raw


def parse_votes(md_file: str) -> pd.DataFrame:
    user_link_re = re.compile(r"\[([^\]]+)\]\(https://vk\.com/[^)]+\)")
    number_re = re.compile(r"([1-9][0-9]?)")
    show_votes_marker = "Показать список оценивших"

    records = []
    with open(md_file, 'r', encoding='utf-8') as f:
        lines = [line.rstrip('\n') for line in f]

    i = 0
    while i < len(lines):
        line = lines[i].strip()
        user_match = user_link_re.search(line)
        if user_match:
            username = user_match.group(1)
            # ищем комментарий
            j = i + 1
            while j < len(lines) and (not lines[j].strip() or lines[j].startswith('![') or lines[j].startswith('[](')):
                j += 1
            if j >= len(lines):
                i += 1
                continue
            comment = lines[j].strip()
            num_match = number_re.search(comment)
            if not num_match:
                i = j + 1
                continue
            participant = int(num_match.group(1))
            # ищем дату
            raw_date = ''
            for k in range(j + 1, len(lines)):
                if show_votes_marker in lines[k]:
                    idx = k + 2
                    if idx < len(lines):
                        m = re.match(r"\[([^]]+)\]\(", lines[idx].strip())
                        if m:
                            raw_date = m.group(1)
                    break
            records.append({'Имя пользователя': username,
                            'Номер участника': participant,
                            'Дата голосования': raw_date})
            i = j + 1
        else:
            i += 1

    df = pd.DataFrame(records)
    df['Дата и время (Excel)'] = df['Дата голосования'].apply(parse_date)
    # добавляем отдельный столбец с датой для Excel
    df['Дата'] = df['Дата и время (Excel)'].apply(lambda x: x.split(' ')[0] if isinstance(x, str) and ' ' in x else x)
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
    parser = argparse.ArgumentParser(description='Подсчет голосов из выгрузки комментариев ВКонтакте')
    parser.add_argument('input_md', help='Путь к .md файлу с комментариями')
    parser.add_argument('--excel', action='store_true', help='Сохранить результаты в Excel (results.xlsx)')
    args = parser.parse_args()

    print('Парсинг голосов...')
    df = parse_votes(args.input_md)

    print('\nТаблица №1 — Сводная по всем голосам:')
    print(summarize_all(df).to_string(index=False))

    print('\nТаблица №2 — Детализация по пользователям:')
    print(detail_by_user(df).to_string(index=False))

    print('\nТаблица №3 — Сводная по уникальным голосам:')
    print(summarize_unique(df).to_string(index=False))

    if args.excel:
        with pd.ExcelWriter('results.xlsx') as writer:
            summarize_all(df).to_excel(writer, sheet_name='Все голоса', index=False)
            detail_by_user(df).to_excel(writer, sheet_name='Детализация', index=False)
            summarize_unique(df).to_excel(writer, sheet_name='Уникальные голоса', index=False)
        print('\nРезультаты сохранены в results.xlsx')

if __name__ == '__main__':
    main()
