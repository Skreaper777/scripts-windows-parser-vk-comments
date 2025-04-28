import re
import argparse
import pandas as pd


def parse_votes(md_file: str):
    """
    Парсит .md файл с комментариями из ВКонтакте и извлекает голоса пользователей.
    Возвращает DataFrame с колонками: Имя пользователя, Номер участника, Дата голосования.
    """
    # Регулярные выражения
    user_link_re = re.compile(r"\[([^\]]+)\]\(https://vk\.com/[^)]+\)")  # ссылка с непустым текстом
    number_re = re.compile(r"([1-9][0-9]?)")  # число от 1 до 99
    show_votes_marker = "Показать список оценивших"

    records = []
    # читаем все строки
    with open(md_file, 'r', encoding='utf-8') as f:
        lines = [line.rstrip('\n') for line in f]

    i = 0
    while i < len(lines):
        line = lines[i].strip()
        # ищем имя пользователя
        user_match = user_link_re.search(line)
        if user_match:
            username = user_match.group(1)
            # следующая строка с комментарием (пропускаем пустые, картинки и пустые ссылки)
            j = i + 1
            while j < len(lines) and (
                not lines[j].strip() or
                lines[j].startswith('![') or
                lines[j].startswith('[](')
            ):
                j += 1
            if j >= len(lines):
                i += 1
                continue
            comment = lines[j].strip()
            # ищем номер участника внутри комментария
            num_match = number_re.search(comment)
            if not num_match:
                i = j + 1
                continue
            participant = int(num_match.group(1))

            # ищем дату голосования: через одну строку после маркера
            date = ""
            for k in range(j + 1, len(lines)):
                if show_votes_marker in lines[k]:
                    idx = k + 2
                    if idx < len(lines):
                        line_date = lines[idx].strip()
                        m = re.match(r"\[([^]]+)\]\(", line_date)
                        if m:
                            date = m.group(1)
                    break

            records.append({
                'Имя пользователя': username,
                'Номер участника': participant,
                'Дата голосования': date
            })
            # продолжаем после обработанного комментария
            i = j + 1
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
    return df.copy().reset_index(drop=True)


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
    parser = argparse.ArgumentParser(
        description='Подсчет голосов из выгрузки комментариев ВКонтакте'
    )
    parser.add_argument('input_md', help='Путь к .md файлу с комментариями')
    parser.add_argument('--excel', action='store_true', help='Сохранить результаты в Excel (results.xlsx)')
    args = parser.parse_args()

    print('Парсинг голосов...')
    df = parse_votes(args.input_md)

    print('\nТаблица №1 — Сводная по всем голосам:')
    table1 = summarize_all(df)
    print(table1.to_string(index=False))

    print('\nТаблица №2 — Детализация по пользователям:')
    table2 = detail_by_user(df)
    print(table2.to_string(index=False))

    print('\nТаблица №3 — Сводная по уникальным голосам:')
    table3 = summarize_unique(df)
    print(table3.to_string(index=False))

    if args.excel:
        with pd.ExcelWriter('results.xlsx') as writer:
            table1.to_excel(writer, sheet_name='Все голоса', index=False)
            table2.to_excel(writer, sheet_name='Детализация', index=False)
            table3.to_excel(writer, sheet_name='Уникальные голоса', index=False)
        print('\nРезультаты сохранены в results.xlsx')


if __name__ == '__main__':
    main()
