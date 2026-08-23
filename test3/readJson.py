import json
from importlib.metadata import requires

"""
with open("data.json", "r", encoding="utf-8") as f:
    data = json.load(f)
"""

"""
            {"type": "category", "text": "Основная информация о родителях"},
            {"type": "input", "label": "Прописка родителей/родителя", "placeholder": "Введите данные"},
            {"type": "input", "label": "Адрес родителей", "placeholder": "Введите данные"},

            {"type": "category", "text": "📍 Адрес"},
            {"type": "combo", "label": "Страна:", "items": ["Россия", "Беларусь", "Казахстан", "Другая"]},
            {"type": "input", "label": "Город:", "placeholder": "Введите город"},

            {"type": "category", "text": "⚙️ Дополнительно"},
            {"type": "checkbox", "label": "Подписаться на новости"},
            {"type": "checkbox", "label": "Согласен с условиями"},

            {"type": "button", "text": "✅ Сохранить данные"}  # кнопка в самом низу
"""

"""
Пример как надо сделать:

EXTERNAL_LISTS = [
    {"type": "input", "label": "Прописка родителей/родителя", "value": "город", "count": "город", "placeholder": "Введите"},
    {"type": "input", "label": "Адрес родителей", "value": "г. Москва, ул. Пушкина 10", "count": "г. Москва"},
    {"type": "checkbox", "label": "Использовать для названия файла", "value": True, "count": True},
    {"type": "combo", "label": "Соц - демографические", "value": "многодетная семья", "count": "многодетная семья"},
]
"""

def add_dict(distionary, key, value, nameCat):
    if key not in distionary:
        distionary[key] = [nameCat]

    distionary[key].append(value)

def read():
    with open("data.json", "r", encoding="utf-8") as f:
        data = json.load(f)

    res = {}

    for item in data:
        if "data" in item:
            # резервное создание переменных
            name = ''
            string = ''
            flags = ''
            category = ''
            numCat = ''
            required = ''
            once = ''
            critical = ''
            lists = ''
            count = ''
            countFlag = ''

            # Проверка на наличие ключей
            if "name" in item["data"]:
                name = item["data"]["name"]
            if "string" in item["data"]:
                string = item["data"]["string"]
            if "flags" in item["data"]:
                flags = item["data"]["flags"]
            if "category" in item["data"]:
                category = item["data"]["category"]
            if "numCat" in item["data"]:
                numCat = item["data"]["numCat"]
            if "required" in item["data"]:
                required = item["data"]["required"]
            if "once" in item["data"]:
                once = item["data"]["once"]
            if "critical" in item["data"]:
                critical = item["data"]["critical"]
            if "list" in item["data"]:
                lists = item["data"]["list"]
            if "count" in item["data"]:
                count = item["data"]["count"]
            if "countFlag" in item["data"]:
                countFlag = item["data"]["countFlag"]

            #print(f"{name}\n{string}\n{flags}\n{category}\n{numCat}\n{required}\n{once}\n{critical}\n{lists}\n{count}\n{countFlag}")
            add_dict(
                distionary=res, key=str(numCat),
                           value={"name":name, "string":string, "flags":flags, "required":required, "once":once, "critical":critical, "list":lists, "count":count, "countFlag":countFlag},
                           nameCat=str(category)
            )

    return res
    pass

print(read())

"""
for item in data:
    if "metadata" in item:
        print(f"Мета: {item['metadata']}")
    if "data" in item:
        print(f"Данные: {item['data']}")
"""

"""
Пример того что лежит в json файле
(между прочим рил данные)
"data": {
            "name": "кемВыявлена",
            "string": "Кем выявлена категория",
            "flags": "rdl",
            "category": "Категории семьи",
            "numCat": 1,
            "required": false,
            "once": false,
            "critical": true,
            "list": [
                " ",
                "ЦСО",
                "СОиП",
                "ОП",
                "УО",
                "КДН",
                "ЗП"
            ],
            "count": "",
            "countFlag": false
        }
"""
