import json

with open("standards.json", "r", encoding="utf-8") as read_file:
    data = json.load(read_file)

# Проход по каждому элементу в данных
for key, value in data.items():
    # print(value["names"])
    # Удаление пустых строк из списка имен и преобразование в множество (set)
    names_set = set(filter(lambda x: x != "", value["names"]))
    # Преобразование множества обратно в список
    value["names"] = list(names_set)

    # Удаление пустых строк из списка артикулов и преобразование в множество (set)
    articles_set = set(filter(lambda x: x != "", value["articles"]))
    # Преобразование множества обратно в список
    value["articles"] = list(articles_set)

    # Удаление строк из списка артикулов, совпадающих со строкой "standart_articul"
    value["articles"] = [article for article in value["articles"] if article != value["standard_article"]]

with open("standards.json", "w", encoding="utf-8") as write_file:
    json.dump(data, write_file, ensure_ascii=False, indent=4)
