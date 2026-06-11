import pandas as pd
import difflib
import json


def find_top_matches(name, references):
    try:
        matches = difflib.get_close_matches(name, references, n=15)
        return matches
    except Exception as error:
        return []


def strip_strings_in_dataframes(dataframes_dict):
    for df_name, df in dataframes_dict.items():
        dataframes_dict[df_name] = df.apply(lambda x: x.str.strip() if x.dtype == 'object' else x)
    return dataframes_dict


def check_column_names(dataframes_dict):
    first_dataframe = list(dataframes_dict.values())[0]
    common_column_names = set(first_dataframe.columns)
    for i, (df_name, df) in enumerate(dataframes_dict.items()):
        different_column_names = set(df.columns) - common_column_names
        missing_column_names = common_column_names - set(df.columns)
        if different_column_names:
            raise Exception(f"Таблица <b>[{i + 1}]'{df_name}'</b> отличается от первой:\n"
                            f"Столбцы отсутствующие в текущей таблице: <b>{different_column_names}</b>")
        if missing_column_names:
            raise Exception(f"Таблица <b>[{i + 1}]'{df_name}'</b> отличается от первой:\n"
                            f"Столбцы отсутствующие в первой таблице: <b>{missing_column_names}</b>")


def dict_to_excel(dataframes_dict, excel_file_path):
    with pd.ExcelWriter(excel_file_path) as writer:
        for sheet_name, df in dataframes_dict.items():
            output = df
            #output = output.assign(Цена=None)
            #output = output.assign(Сумма=None)
            output.to_excel(writer, sheet_name=sheet_name, index=False)


def measure_check(measure_value):
    measures = {
        'шт': ['шт.', 'штук', 'штука'],  # компл. нужно учитывать?
        'м': ['м.', 'метр', 'meter'],
        'м2': ['м. в кв.'],
        'кг': ['кг.', 'килограмм']
    }
    measure_value = str(measure_value).lower()
    replacement = None
    for key, values in measures.items():
        if measure_value == key:
            replacement = key
            break
        elif measure_value in values:
            replacement = key
            break
    return replacement


def save_new_object(data, standard_name, standard_article, measure, name, article):
    if name:
        name_list = [name]
    else:
        name_list = []

    if article:
        article_list = [article]
    else:
        article_list = []

    data[standard_name] = {
        "standard_article": standard_article,
        "measure": measure,
        "names": name_list,
        "articles": article_list
    }


def save_non_standard_name(data, name, standard_name):
    standard = data[standard_name]
    if pd.notna(name) and name not in standard['names'] and name != standard_name:
        data[standard_name]['names'].append(name)


def save_non_standard_article(data, article, standard_name):
    standard = data[standard_name]
    if pd.notna(article) and article not in standard['articles'] and article != standard['standard_article']:
        data[standard_name]['articles'].append(article)


def clean_suggestions(suggestions):
    clear_suggestions = []
    values = []
    for suggestion in suggestions:
        if f"[{suggestion[1]} | {suggestion[0]}]" not in values:
            values.append(f"[{suggestion[1]} | {suggestion[0]}]")
            clear_suggestions.append(suggestion)
        if len(clear_suggestions) == 15:
            break
    return clear_suggestions


# Нужна ли проверка единиц измерения на соответствие с базой? (Я думаю - нет)
# Нужно ли прорабатывать случай, когда в таблице нет колонки с артикулами? (Я думаю - нет)
# Что делать, если эталонный артикул пустой?
# Что делать, если не нужно добавлять?
# Нужно ли анализировать колонку "Тип, марка, обозначение документа"?
# Одно и тоже возможное наименование может относиться к нескольким эталоном, а артикул нет?
# Что делать если эталонного артикула нет, а в таблице какой-то есть?
# Компл. нужно в штуки учитывать?

def get_suggestions(dfs, main_columns, add_info_columns):
    # Настраиваемые параметры

    articles_column = main_columns[0]
    names_column = main_columns[1]
    quantity_column = main_columns[2]
    measure_column = main_columns[3]

    columns = [articles_column, names_column, measure_column]
    columns.extend(add_info_columns)

    with open("standards.json", "r", encoding="utf-8") as read_file:
        data = json.load(read_file)

    suggestions_list = []
    warn_dict = {}
    warn_index = 0

    for df_name, df in dfs.items():
        for index, row in df.iterrows():
            measure_value = row.iloc[measure_column]
            if measure_check(measure_value) is not None:
                row.iloc[measure_column] = measure_check(measure_value)
            #else:
            #    print(f"Незнакомая единица измерения: [{row.iloc[measure_column]}]")

        input_data = df.iloc[:, columns]

        for index, row in input_data.iterrows():


            warn_dict[warn_index] = ""
            article = row.iloc[0]  # Получаем значение столбца 'articles_column' из текущей строки
            name = row.iloc[1]  # Получаем значение столбца 'names_column' из текущей строки
            measure = row.iloc[2]  # Получаем значение столбца 'measure_column' из текущей строки

            if measure_check(measure) is not None:
                row.iloc[2] = measure_check(measure)

            extended_name = name
            i = 0
            for column in add_info_columns:
                if str(row.iloc[3 + i]) != "nan":
                    extended_name = extended_name + " , " + str(row.iloc[3 + i])
                i = i + 1

            suggestions = []
            certain_suggestion_count = 0  # счётчик "уверенных" предложений

            for standard_name, standard in data.items():

                if standard_name == name:
                    if standard['standard_article'] == article:
                        description = f"Наименование <b>{name}</b> и артикул <b>{article}</b> уже эталонные"
                        suggestions.append([name, article, standard['measure'], 1, description])
                    elif standard['standard_article'] != "":
                        description = (
                            f"Наименование <b>{name}</b> - уже эталонное, и ему соответствует другой эталонный артикул "
                            f"<b>{standard['standard_article']}</b>, а не <b>{article}</b>")
                        suggestions.append([name, standard['standard_article'], standard['measure'], 2, description])
                    else:
                        description = f"Наименование <b>{name}</b> - уже эталонное, а эталонного артикула - нет"
                        suggestions.extend([name, "", standard['measure'], 3, description])

                    certain_suggestion_count = certain_suggestion_count + 1

                elif standard['standard_article'] == article:
                    description = f"Артикул <b>{article}</b> для <b>{name}</b> совпадает с эталонным для <b>{standard_name}</b>"
                    suggestions.append([standard_name, article, standard['measure'], 4, description])
                    certain_suggestion_count = certain_suggestion_count + 1

                elif article in standard['articles']:
                    if name in standard['names']:
                        description = (f"Наименование <b>{name}</b> и артикул <b>{article}</b> "
                                       f"присутствуют в базе данных как ранее заменённые для "
                                       f"эталона <b>{standard_name}</b> c артиклем <b>{standard['standard_article']}</b>")
                        suggestions.append([standard_name, standard['standard_article'], standard['measure'], 5, description])
                        certain_suggestion_count = certain_suggestion_count + 1
                    else:
                        description = (
                            f"Артикул <b>{article}</b> присутствует в базе данных как ранее заменённый для "
                            f"эталона <b>{standard_name}</b> c артиклем <b>{standard['standard_article']}</b>")
                        suggestions.append([standard_name, standard['standard_article'], standard['measure'], 6, description])
                        certain_suggestion_count = certain_suggestion_count + 1

                elif extended_name in standard['names']:
                    description = (f"Расширенное наименование <b>{extended_name}</b> присутствует в базе данных "
                                   f"как ранее заменённое для "
                                   f"эталона <b>{standard_name}</b> c артиклем <b>{standard['standard_article']}</b>")
                    suggestions.append([standard_name, standard['standard_article'], standard['measure'], 7, description])
                    certain_suggestion_count = certain_suggestion_count + 1

                elif name in standard['names']:
                    description = (f"Наименование <b>{name}</b> присутствует в базе данных как ранее заменённое для "
                                   f"эталона <b>{standard_name}</b> c артиклем <b>{standard['standard_article']}</b>")
                    suggestions.append([standard_name, standard['standard_article'], standard['measure'], 7, description])
                    certain_suggestion_count = certain_suggestion_count + 1

            if certain_suggestion_count > 1:
                warn_dict[warn_index] = (f"Предупреждение: [<b>{df_name}</b>] Строка <b>{index+1}</b>: "
                                         f"Больше 1 уверенных предложений на строчку (<b>{certain_suggestion_count}</b>)")

            standard_names = list(data.keys())
            matches = find_top_matches(extended_name, standard_names)

            for string in find_top_matches(name, standard_names):
                if string not in matches:
                    matches.append(string)

            if matches:
                for string in matches:
                    standard = data[string]
                    description = f"<b>{extended_name}</b> схоже на <b>{string}</b>"
                    suggestions.append([string, standard['standard_article'], standard['measure'], 8, description])

            if suggestions:
                suggestions = sorted(suggestions, key=lambda x: x[3])
                suggestions = clean_suggestions(suggestions)

            else:
                suggestions.append("")
                warn_dict[warn_index] = f"<b>{article} | {extended_name}</b>: Замен не нашлось"
            suggestions_list.append(suggestions)
            warn_index = warn_index + 1

    return suggestions_list, warn_dict
