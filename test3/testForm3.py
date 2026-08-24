import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt

# ВНЕШНИЙ СПИСОК - ОСНОВНОЙ ИСТОЧНИК ДАННЫХ
EXTERNAL_LISTS = [
    {"type": "category", "text": "Основная информация о родителях"},
    {"type": "input", "label": "Прописка родителей/родителя", "value": "город", "count": "город"},
    {"type": "input", "label": "Адрес родителей", "value": "г. Москва, ул. Пушкина 10", "count": "г. Москва"},
    {"type": "category", "text": "Социальные факторы"},
    {"type": "category", "text": "Основная информация о родителях"},
    {"type": "checkbox", "label": "Использовать для названия файла", "value": True, "count": True},
    {"type": "combo", "label": "Соц - демографические",
     "items": [" ", "одинокая мать", "замещающая семья", "многодетная семья", "разведённые родители"],
     "value": "многодетная семья", "count": "многодетная семья"},
    {"type": "category", "text": "Основная информация о родителях"},
    {"type": "category", "text": "Дополнительная информация"},
    {"type": "button", "text": "✅ Сохранить данные"}
]


class DynamicForm(QMainWindow):
    def __init__(self, external_lists=None):
        super().__init__()
        self.setWindowTitle("Динамическая форма")
        self.setGeometry(100, 100, 800, 600)

        # ЕСЛИ ПЕРЕДАН ВНЕШНИЙ СПИСОК - ИСПОЛЬЗУЕМ ЕГО
        if external_lists:
            self.form_data = external_lists.copy()
        else:
            # ИНАЧЕ ДЕФОЛТНЫЙ СПИСОК С ЗАГОЛОВКАМИ
            self.form_data = [
                {"type": "category", "text": "Основная информация о родителях"},
                {"type": "input", "label": "Прописка родителей/родителя", "placeholder": "Введите данные"},
                {"type": "input", "label": "Адрес родителей", "placeholder": "Введите данные"},
                {"type": "category", "text": "Социальные факторы"},
                {"type": "checkbox", "label": "Использовать для названия файла"},
                {"type": "combo", "label": "Соц - демографические",
                 "items": [" ", "одинокая мать", "замещающая семья", "многодетная семья", "разведённые родители"]},
                {"type": "category", "text": "Дополнительная информация"},
                {"type": "button", "text": "✅ Сохранить данные"}
            ]

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)

        canvas = QWidget()
        self.layout = QVBoxLayout(canvas)
        self.layout.setSpacing(10)
        self.layout.setContentsMargins(10, 10, 10, 10)

        self.build_form()

        scroll.setWidget(canvas)
        central_widget_layout = QVBoxLayout(central_widget)
        central_widget_layout.addWidget(scroll)
        central_widget_layout.setContentsMargins(0, 0, 0, 0)

    def build_form(self):
        """СТРОИМ ФОРМУ ИЗ self.form_data (КОТОРЫЙ МОЖЕТ БЫТЬ ИЗ EXTERNAL)"""
        self.clear_layout(self.layout)
        self.input_fields = []

        for item in self.form_data:
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

            # ---------- ПОЛЕ ВВОДА ----------
            elif item_type == "input":
                container = QWidget()
                h_layout = QHBoxLayout(container)
                h_layout.setContentsMargins(0, 0, 0, 0)

                lbl = QLabel(item["label"])
                lbl.setFixedWidth(200)
                lbl.setStyleSheet("font-weight: 500;")

                inp = QLineEdit()

                # ЕСЛИ ЕСТЬ placeholder
                if "placeholder" in item:
                    inp.setPlaceholderText(item["placeholder"])

                # ПОДСТАВЛЯЕМ ЗНАЧЕНИЕ
                if "value" in item and item["value"]:
                    inp.setText(str(item["value"]))
                elif "count" in item and item["count"]:
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

                inp.field_id = item["label"]
                inp.field_type = "input"
                self.input_fields.append(inp)

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
                if "items" in item:
                    combo.addItems(item["items"])
                else:
                    combo.addItems([" "])

                selected_value = None
                if "value" in item and item["value"]:
                    selected_value = item["value"]
                elif "count" in item and item["count"]:
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

            # ---------- ЧЕКБОКС ----------
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
                btn.clicked.connect(self.save_data)
                self.layout.addWidget(btn)

        self.layout.addStretch()

    def clear_layout(self, layout):
        if layout is not None:
            while layout.count():
                item = layout.takeAt(0)
                widget = item.widget()
                if widget is not None:
                    widget.deleteLater()
                else:
                    self.clear_layout(item.layout())

    def save_data(self):
        print("\n" + "=" * 50)
        print("📊 СОБРАННЫЕ ДАННЫЕ:")
        print("=" * 50)

        for widget in self.input_fields:
            field_name = getattr(widget, 'field_id', 'Неизвестное поле')
            field_type = getattr(widget, 'field_type', 'unknown')

            for item in self.form_data:
                if item.get("label") == field_name and item.get("type") == field_type:
                    if isinstance(widget, QLineEdit):
                        value = widget.text()
                        item["value"] = value
                        item["count"] = value if value else None
                        print(f"{field_name:20} → {value if value else '[пусто]'}")

                    elif isinstance(widget, QComboBox):
                        value = widget.currentText()
                        item["value"] = value
                        item["count"] = value if value and value != " " else None
                        print(f"{field_name:20} → {value}")

                    elif isinstance(widget, QCheckBox):
                        value = widget.isChecked()
                        item["value"] = value
                        item["count"] = True if value else False
                        print(f"{field_name:20} → {'✅ Да' if value else '❌ Нет'}")
                    break

        print("=" * 50)

        reply = QMessageBox.question(
            self,
            'Подтверждение',
            'Сохранить изменения?',
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            QMessageBox.information(self, 'Сохранено', 'Данные сохранены!')
            print("\n✅ ОБНОВЛЕННЫЙ self.form_data:")
            for item in self.form_data:
                if "value" in item or "count" in item:
                    print(f"  {item}")
            return self.form_data
        else:
            print("\n❌ Сохранение отменено")
            return None

    def closeEvent(self, event):
        reply = QMessageBox.question(
            self,
            'Подтверждение',
            'Вы действительно хотите выйти?',
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()


if __name__ == "__main__":
    app = QApplication(sys.argv)

    # МОЖНО ИСПОЛЬЗОВАТЬ С ВНЕШНИМ СПИСКОМ (С ЗАГОЛОВКАМИ)
    window = DynamicForm(EXTERNAL_LISTS)
    window.show()

    # ИЛИ БЕЗ НЕГО (ТОГДА ИСПОЛЬЗУЕТ ДЕФОЛТНЫЙ СПИСОК ТОЖЕ С ЗАГОЛОВКАМИ)
    # window = DynamicForm()
    # window.show()

    sys.exit(app.exec_())