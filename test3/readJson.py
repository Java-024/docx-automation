import json

with open("data.json", "r", encoding="utf-8") as f:
    data = json.load(f)

#print(data)

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


for item in data:
    if "metadata" in item:
        print(f"Мета: {item['metadata']}")
    if "data" in item:
        print(f"Данные: {item['data']}")