"""
пример

lists = [
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
        ]
"""

import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt


class DynamicForm(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Динамическая форма")
        self.setGeometry(100, 100, 800, 600)  # x, y, ширина, высота

        # ---------- 1. ГЛАВНЫЙ КОНТЕЙНЕР ----------
        # Создаём центральный виджет (обязательно для QMainWindow)
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # ---------- 2. ОБЛАСТЬ ПРОКРУТКИ ----------
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)  # чтобы подстраивалось под содержимое

        # ---------- 3. ХОЛСТ ДЛЯ ВИДЖЕТОВ ----------
        canvas = QWidget()
        self.layout = QVBoxLayout(canvas)  # вертикальное расположение
        self.layout.setSpacing(10)  # расстояние между элементами
        self.layout.setContentsMargins(10, 10, 10, 10)  # отступы от краев

        # ---------- 4. ДАННЫЕ ДЛЯ ФОРМЫ (СПИСОК) ----------
        self.form_data = [
            {"type": "category", "text": "Основная информация о родителях"},
            {"type": "input", "label": "Прописка родителей/родителя", "placeholder": "Введите данные"},
            {"type": "input", "label": "Адрес родителей", "placeholder": "Введите данные"},
            {"type": "checkbox", "label": "Использовать для названия файла"},

            {"type": "category", "text": "Социальные факторы"},
            {"type": "combo", "label": "Соц - демографические", "items": [" ", "одинокая мать",
                                                                          "замещающая семья", "многодетная семья",
                                                                          "разведённые родители"]},
            #{"type": "input", "label": "Город:", "placeholder": "Введите город"},

            #{"type": "category", "text": "⚙️ Дополнительно"},
            #{"type": "checkbox", "label": "Подписаться на новости"},
            #{"type": "checkbox", "label": "Согласен с условиями"},

            {"type": "button", "text": "✅ Сохранить данные"}  # кнопка в самом низу
        ]

        # ---------- 5. ПОСТРОЕНИЕ ФОРМЫ ----------
        self.build_form(self.form_data)

        # ---------- 6. СБОРКА ВСЕХ КОМПОНЕНТОВ ----------
        scroll.setWidget(canvas)
        central_widget_layout = QVBoxLayout(central_widget)
        central_widget_layout.addWidget(scroll)
        central_widget_layout.setContentsMargins(0, 0, 0, 0)  # чтобы скролл был вплотную

    # -------- ФУНКЦИЯ ПОСТРОЕНИЯ ФОРМЫ (ГЛАВНАЯ ЛОГИКА) --------
    def build_form(self, data):
        # Очищаем старые виджеты (если форма перестраивается)
        self.clear_layout(self.layout)

        # Создаем список, чтобы потом собирать данные со всех полей
        self.input_fields = []  # здесь будем хранить все поля ввода

        # Проходим по каждому элементу в списке
        for item in data:
            item_type = item.get("type")

            # ---------- КАТЕГОРИЯ (ЗАГОЛОВОК) ----------
            if item_type == "category":
                label = QLabel(item["text"])
                label.setStyleSheet("""
                    font-size: 14px;
                    font-weight: bold;
                    color: #2c3e50;
                    margin-top: 15px;
                    padding: 5px;
                    background-color: #ecf0f1;
                    border-radius: 3px;
                """)
                self.layout.addWidget(label)

            # ---------- ПОЛЕ ВВОДА ТЕКСТА ----------
            elif item_type == "input":
                # Создаем горизонтальный контейнер для label + поле ввода
                container = QWidget()
                h_layout = QHBoxLayout(container)
                h_layout.setContentsMargins(0, 0, 0, 0)

                # Создаем label
                lbl = QLabel(item["label"])
                lbl.setFixedWidth(200)  # фиксированная ширина для ровности
                lbl.setStyleSheet("font-weight: 500;")

                # Создаем поле ввода
                inp = QLineEdit()
                inp.setPlaceholderText(item.get("placeholder", ""))
                inp.setStyleSheet("""
                    QLineEdit {
                        padding: 5px;
                        border: 1px solid #bdc3c7;
                        border-radius: 3px;
                    }
                    QLineEdit:focus {
                        border: 1px solid #3498db;
                    }
                """)

                # Сохраняем виджет для сбора данных (запоминаем по label)
                inp.field_id = item["label"]
                self.input_fields.append(inp)

                # Добавляем в контейнер
                h_layout.addWidget(lbl)
                h_layout.addWidget(inp)
                self.layout.addWidget(container)

            # ---------- ВЫПАДАЮЩИЙ СПИСОК ----------
            elif item_type == "combo":
                container = QWidget()
                h_layout = QHBoxLayout(container)
                h_layout.setContentsMargins(0, 0, 0, 0)

                lbl = QLabel(item["label"])
                lbl.setFixedWidth(100)
                lbl.setStyleSheet("font-weight: 500;")

                combo = QComboBox()
                combo.addItems(item["items"])
                combo.setStyleSheet("""
                    QComboBox {
                        padding: 5px;
                        border: 1px solid #bdc3c7;
                        border-radius: 3px;
                    }
                    QComboBox::drop-down {
                        border: none;
                    }
                """)

                combo.field_id = item["label"]
                self.input_fields.append(combo)

                h_layout.addWidget(lbl)
                h_layout.addWidget(combo)
                self.layout.addWidget(container)

            # ---------- ГАЛОЧКА (ЧЕКБОКС) ----------
            elif item_type == "checkbox":
                chk = QCheckBox(item["label"])
                chk.setStyleSheet("""
                    QCheckBox {
                        spacing: 5px;
                        font-weight: 500;
                    }
                    QCheckBox::indicator {
                        width: 18px;
                        height: 18px;
                    }
                """)
                chk.field_id = item["label"]
                self.input_fields.append(chk)
                self.layout.addWidget(chk)

            # ---------- КНОПКА ----------
            elif item_type == "button":
                btn = QPushButton(item["text"])
                btn.setStyleSheet("""
                    QPushButton {
                        background-color: #3498db;
                        color: white;
                        border: none;
                        padding: 10px;
                        border-radius: 5px;
                        font-weight: bold;
                        font-size: 14px;
                    }
                    QPushButton:hover {
                        background-color: #2980b9;
                    }
                    QPushButton:pressed {
                        background-color: #21618c;
                    }
                """)
                btn.clicked.connect(self.save_data)  # привязываем функцию сохранения
                self.layout.addWidget(btn)

        # Добавляем распорку в конец (чтобы элементы были сверху, а не по центру)
        self.layout.addStretch()

    # -------- ФУНКЦИЯ ОЧИСТКИ ЛЕЙАУТА --------
    def clear_layout(self, layout):
        if layout is not None:
            while layout.count():
                item = layout.takeAt(0)
                widget = item.widget()
                if widget is not None:
                    widget.deleteLater()
                else:
                    self.clear_layout(item.layout())

    # -------- ФУНКЦИЯ СБОРА ДАННЫХ (ПРИ НАЖАТИИ КНОПКИ) --------
    def save_data(self):
        print("\n" + "=" * 50)
        print("📊 СОБРАННЫЕ ДАННЫЕ:")
        print("=" * 50)

        for widget in self.input_fields:
            field_name = getattr(widget, 'field_id', 'Неизвестное поле')

            # Проверяем тип виджета и берем данные
            if isinstance(widget, QLineEdit):
                value = widget.text() if widget.text() else "[пусто]"
            elif isinstance(widget, QComboBox):
                value = widget.currentText()
            elif isinstance(widget, QCheckBox):
                value = "✅ Да" if widget.isChecked() else "❌ Нет"
            else:
                value = "[неизвестный тип]"

            print(f"{field_name:20} → {value}")
        print("=" * 50 + "\n")


# -------- ЗАПУСК ПРОГРАММЫ --------
if __name__ == "__main__":
    #print(lists)
    app = QApplication(sys.argv)
    window = DynamicForm()
    window.show()
    sys.exit(app.exec_())