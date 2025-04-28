import re
import argparse
import pandas as pd


def parse_votes(md_file: str):
    """
    Парсит .md файл с комментариями из ВКонтакте и извлекает голоса пользователей.
    Возвращает DataFrame с колонками: Имя пользователя, Номер участника, Дата голосования.
    """
    # Регулярки
    user_link_re = re.compile(r"\[([^\]]+)\]\(https://vk\.com/[^)]+\)")  # ссылка с непустым текстом
    number_re = re.compile(r"([1-9][0-9]?)")  # число от 1 до 99
    date_re = re.compile(r"\[(\d{1,2} (?:янв|фев|мар|апр|май|июн|июл|авг|сен|окт|ноя|дек) в \d{1,2}:\d{2})\]", re.IGNORECASE)
    show_votes_marker = "Показать список оценивших"

    records = []
    with open(md_file, 'r', encoding='utf-8') as f:
        lines = [line.strip() for line in f]

    i = 0
    while i < len(lines):
        line = lines[i]
        # Ищем строку с пользователем
        if user_link_re.search(line) and not line.startswith('!['):
            username = user_link_re.search(line).group(1)
            # Собираем блок комментария до следующего пользователя
            block = []
            i += 1
            while i < len(lines) and not (user_link_re.search(lines[i]) and not lines[i].startswith('![')):
                block.append(lines[i])
                i += 1

            # Извлекаем номер участника
            participant = None
            for b in block:
                if not b or b.startswith('![') or b.startswith('[]('):
                    continue
                num_match = number_re.search(b)
                if num_match:
                    participant = int(num_match.group(1))
                    break

            if participant is None:
                # некорректный голос
                continue

            # Извлекаем дату голосования
            date = ""
            # находим маркер, затем ищем дату в следующих строках
            for idx, b in enumerate(block):
                if show_votes_marker in b:
                    for later in block[idx+1:]:
                        date_match = date_re.search(later)
                        if date_match:
                            date = date_match.group(1)
                            break
                    break

            records.append({
                'Имя пользователя': username,
                'Номер участника': participant,
                'Дата голосования': date
            })
        else:
            i += 1

    df = pd.DataFrame(records)
    return df


def summarize_all(df: pd.DataFrame):
    """
    Таблица №1 — Сводная по номерам участников (все голоса).
    """
    summary = (
        df['Номер участника']
        .value_counts()
        .sort_index()
        .rename_axis('Номер участника')
        .reset_index(name='Количество голосов')
    )
    return summary


def detail_by_user(df: pd.DataFrame):
    """
    Таблица №2 — Детализация по каждому пользователю.
    """
    detail = df.copy().reset_index(drop=True)
    return detail


def summarize_unique(df: pd.DataFrame):
    """
    Таблица №3 — Сводная по номерам участников (уникальные пользователи).
    """
    unique_votes = df.drop_duplicates(subset=['Имя пользователя'], keep='first')
    summary = (
        unique_votes['Номер участника']
        .value_counts()
        .sort_index()
        .rename_axis('Номер участника')
        .reset_index(name='Уникальные голоса')
    )
    return summary


def main():
    parser = argparse.ArgumentParser(description='Подсчет голосов из выгрузки комментариев ВКонтакте')
    parser.add_argument('input_md', help='Путь к .md файлу с комментариями')
    parser.add_argument('--excel', action='store_true', help='Сохранить результаты в Excel (results.xlsx)')
    args = parser.parse_args()

    print('Парсинг голосов...')
    df = parse_votes(args.input_md)

    print('\nТаблица №1 — Сводная по всем голосам:')
    table1 = summarize_all(df)
    print(table1.to_markdown(index=False))

    print('\nТаблица №2 — Детализация по пользователям:')
    table2 = detail_by_user(df)
    print(table2.to_markdown(index=False))

    print('\nТаблица №3 — Сводная по уникальным голосам:')
    table3 = summarize_unique(df)
    print(table3.to_markdown(index=False))

    if args.excel:
        with pd.ExcelWriter('results.xlsx') as writer:
            table1.to_excel(writer, sheet_name='Все голоса', index=False)
            table2.to_excel(writer, sheet_name='Детализация', index=False)
            table3.to_excel(writer, sheet_name='Уникальные голоса', index=False)
        print('\nРезультаты сохранены в results.xlsx')


if __name__ == '__main__':
    main()
