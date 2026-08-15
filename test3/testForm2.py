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

# ДОБАВЛЯЕМ: внешний список с начальными значениями
# Если этот список передан, то используем его значения для заполнения полей
EXTERNAL_LISTS = [
    {"type": "input", "label": "Прописка родителей/родителя", "value": "город", "count": "город", "placeholder": "Введите"},
    {"type": "input", "label": "Адрес родителей", "value": "г. Москва, ул. Пушкина 10", "count": "г. Москва"},
    {"type": "checkbox", "label": "Использовать для названия файла", "value": True, "count": True},
    {"type": "combo", "label": "Соц - демографические", "value": "многодетная семья", "count": "многодетная семья"},
]

class DynamicForm(QMainWindow):
    def __init__(self, external_lists=None):
        super().__init__()
        self.setWindowTitle("Динамическая форма")
        self.setGeometry(100, 100, 800, 600)  # x, y, ширина, высота

        # ДОБАВЛЯЕМ: сохраняем внешний список если он есть
        self.external_lists = external_lists

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

        # ДОБАВЛЯЕМ: применяем внешние значения к form_data
        if self.external_lists:
            self.apply_external_values()

        # ---------- 5. ПОСТРОЕНИЕ ФОРМЫ ----------
        self.build_form(self.form_data)

        # ---------- 6. СБОРКА ВСЕХ КОМПОНЕНТОВ ----------
        scroll.setWidget(canvas)
        central_widget_layout = QVBoxLayout(central_widget)
        central_widget_layout.addWidget(scroll)
        central_widget_layout.setContentsMargins(0, 0, 0, 0)  # чтобы скролл был вплотную

    # ДОБАВЛЯЕМ: функция применения внешних значений
    def apply_external_values(self):
        """Применяет значения из внешнего списка к form_data"""
        for ext_item in self.external_lists:
            ext_label = ext_item.get("label")
            ext_type = ext_item.get("type")
            ext_value = ext_item.get("value")
            ext_count = ext_item.get("count")

            # Ищем соответствующий элемент в form_data
            for form_item in self.form_data:
                if form_item.get("label") == ext_label and form_item.get("type") == ext_type:
                    # Если есть value - добавляем его
                    if ext_value is not None:
                        form_item["value"] = ext_value
                    # Если есть count - добавляем его
                    if ext_count is not None:
                        form_item["count"] = ext_count
                    break

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

                # ДОБАВЛЯЕМ: если есть value или count - подставляем в поле
                if "value" in item and item["value"]:
                    inp.setText(item["value"])
                elif "count" in item and isinstance(item["count"], str):
                    inp.setText(item["count"])
                elif "count" in item and item["count"] is not None and not isinstance(item["count"], bool):
                    inp.setText(str(item["count"]))

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
                inp.field_type = "input"
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

                # ДОБАВЛЯЕМ: если есть value или count - выбираем нужный элемент
                selected_value = None
                if "value" in item and item["value"]:
                    selected_value = item["value"]
                elif "count" in item and isinstance(item["count"], str) and item["count"]:
                    selected_value = item["count"]

                if selected_value:
                    index = combo.findText(selected_value)
                    if index >= 0:
                        combo.setCurrentIndex(index)

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
                combo.field_type = "combo"
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
                chk.field_type = "checkbox"

                # ДОБАВЛЯЕМ: если есть value или count - ставим галочку
                if "value" in item:
                    if item["value"] is True:
                        chk.setChecked(True)
                    elif item["value"] is False:
                        chk.setChecked(False)
                elif "count" in item:
                    if item["count"] is True:
                        chk.setChecked(True)
                    elif item["count"] is False:
                        chk.setChecked(False)
                # если None или строка - оставляем как есть

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
        # СОЗДАЕМ КОПИЮ lists С ДОБАВЛЕННЫМ count
        lists_copy = []

        print("\n" + "=" * 50)
        print("📊 СОБРАННЫЕ ДАННЫЕ:")
        print("=" * 50)

        for widget in self.input_fields:
            field_name = getattr(widget, 'field_id', 'Неизвестное поле')
            field_type = getattr(widget, 'field_type', 'unknown')

            # Проверяем тип виджета и берем данные
            if isinstance(widget, QLineEdit):
                value = widget.text() if widget.text() else ""
                # ДОБАВЛЯЕМ: создаем словарь для копии
                lists_copy.append({
                    "type": "input",
                    "label": field_name,
                    "value": value,
                    "count": value if value else None  # если есть текст - строка, иначе None
                })
                print(f"{field_name:20} → {value if value else '[пусто]'}")

            elif isinstance(widget, QComboBox):
                value = widget.currentText()
                lists_copy.append({
                    "type": "combo",
                    "label": field_name,
                    "value": value,
                    "count": value if value and value != " " else None
                })
                print(f"{field_name:20} → {value}")

            elif isinstance(widget, QCheckBox):
                value = widget.isChecked()
                # ДОБАВЛЯЕМ: сохраняем count как True/False
                lists_copy.append({
                    "type": "checkbox",
                    "label": field_name,
                    "value": value,
                    "count": True if value else False
                })
                print(f"{field_name:20} → {'✅ Да' if value else '❌ Нет'}")

            else:
                value = "[неизвестный тип]"
                print(f"{field_name:20} → {value}")

        print("=" * 50)

        # ДОБАВЛЯЕМ: ДИАЛОГ ПОДТВЕРЖДЕНИЯ ДЛЯ СОХРАНЕНИЯ
        reply = QMessageBox.question(
            self,
            'Подтверждение',
            'Сохранить изменения?',
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            # СОХРАНЯЕМ В КОПИЮ (можно использовать lists_copy)
            # Здесь можно записать в файл или другую логику
            QMessageBox.information(self, 'Сохранено', 'Данные сохранены!')
            print("\n✅ Данные сохранены в lists_copy:")
            for item in lists_copy:
                print(f"  {item}")
            # Окно НЕ закрывается
        else:
            print("\n❌ Сохранение отменено")
            # Окно НЕ закрывается

        # ДОБАВЛЯЕМ: возвращаем lists_copy для использования в другом месте
        return lists_copy

    # ДОБАВЛЯЕМ: ФУНКЦИЯ ЗАКРЫТИЯ С ПОДТВЕРЖДЕНИЕМ
    def closeEvent(self, event):
        """Переопределяем стандартное закрытие окна"""
        reply = QMessageBox.question(
            self,
            'Подтверждение',
            'Вы действительно хотите выйти?',
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            event.accept()  # закрываем окно
        else:
            event.ignore()  # не закрываем окно


# -------- ЗАПУСК ПРОГРАММЫ --------
if __name__ == "__main__":
    app = QApplication(sys.argv)

    # ДОБАВЛЯЕМ: проверяем, есть ли внешний список
    # Если EXTERNAL_LISTS существует и не пустой - используем его
    # Иначе создаем окно без внешних значений
    try:
        if EXTERNAL_LISTS and len(EXTERNAL_LISTS) > 0:
            window = DynamicForm(EXTERNAL_LISTS)
        else:
            window = DynamicForm()
    except NameError:
        # Если EXTERNAL_LISTS не определен - используем стандартный вариант
        window = DynamicForm()

    window.show()
    sys.exit(app.exec_())