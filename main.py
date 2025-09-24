import GDocument
import sys
from PyQt5.QtWidgets import QApplication, QMessageBox, QComboBox, QCheckBox, QLineEdit, QWidget, QVBoxLayout, QPushButton, QTextEdit, QDialog, QLabel, QHBoxLayout, QDialogButtonBox
import traceback


class GenerateParametersDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Параметры генерации")
        self.setModal(True)
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout()

        def row(label_text, placeholder):
            h = QHBoxLayout()
            h.addWidget(QLabel(label_text))
            edit = QLineEdit()
            edit.setPlaceholderText(placeholder)
            h.addWidget(edit)
            layout.addLayout(h)
            return edit

        # Отдельные поля (не переопределяем одну и ту же переменную!)
        # self.tb_name_edit   = row("Название таблицы с участниками для рассылки:", "TB_NAME")
        self.event_name_edit= row("Краткое название мероприятия (н-р, NPW-2025):", "EVENT_NAME")
        self.event_info_edit= row("Полное название мероприятия (П.п., «в чём?»):", "EVENT_INFO")
        self.date_info_edit = row("Даты проведения (н-р, 7–13 сентября):", "DATE_INFO")
        self.place_info_edit= row("Место проведения (н-р, г. Москва):", "PLACE_INFO")
        self.oferta_link_edit=row("Ссылка на Публичную оферту:", "OFERTA_LINK")

        # Кнопки
        btn_layout = QHBoxLayout()
        btn_ok = QPushButton("OK")
        btn_cancel = QPushButton("Отмена")
        btn_ok.clicked.connect(self.on_ok)
        btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(btn_ok)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)

        self.setLayout(layout)

    def on_ok(self):
        # Простейшая валидация — можно расширить
        if not self.event_name_edit.text().strip():
            QMessageBox.warning(self, "Проверка", "Укажите TB_NAME (путь к таблице).")
            return
        self.accept()

    def get_parameters(self):
        """Возвращает параметры в виде словаря с нужными ключами."""
        return {
            # 'TB_NAME': self.tb_name_edit.text().strip(),
            'EVENT_NAME': self.event_name_edit.text().strip(),
            'EVENT_INFO': self.event_info_edit.text().strip(),
            'DATE_INFO': self.date_info_edit.text().strip(),
            'PLACE_INFO': self.place_info_edit.text().strip(),
            'OFERTA_LINK': self.oferta_link_edit.text().strip(),
        }

class BoolParameterDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Выбор параметра")
        self.setModal(True)
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Текст вопроса
        label = QLabel("Выберите параметры отправки:")
        layout.addWidget(label)
        
        # Кнопки для выбора True/False
        button_layout = QHBoxLayout()
        
        self.true_btn = QPushButton("Тестовое письмо")
        self.true_btn.clicked.connect(self.true_selected)
        button_layout.addWidget(self.true_btn)
        
        self.false_btn = QPushButton("Отправить рассылку")
        self.false_btn.clicked.connect(self.false_selected)
        button_layout.addWidget(self.false_btn)
        
        layout.addLayout(button_layout)
        
        # Кнопки отмены
        button_box = QDialogButtonBox(QDialogButtonBox.Cancel)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
        self.setLayout(layout)
        
        # Изначально параметр не выбран
        self.selected_value = None
    
    def true_selected(self):
        self.selected_value = True
        self.accept()
    
    def false_selected(self):
        self.selected_value = False
        self.accept()

import sys
import os
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QTextEdit, 
                             QPushButton, QDialog)
# from config import Config

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Функции с параметрами и история")
        # self.config = Config()
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        self.history = QTextEdit()
        self.history.setReadOnly(True)
        self.history.setPlaceholderText("История выполнений появится здесь...")
        layout.addWidget(self.history)

        btn1 = QPushButton("Создать шаблоны")
        btn1.clicked.connect(lambda: self.run_and_log(GDocument.create.create_all_templates))
        layout.addWidget(btn1)

        btn2 = QPushButton("Сгенерировать")
        btn2.clicked.connect(self.on_btn2_clicked)
        layout.addWidget(btn2)

        btn3 = QPushButton("Рассылка")
        btn3.clicked.connect(self.on_btn3_clicked)
        layout.addWidget(btn3)

        self.setLayout(layout)

    def on_btn2_clicked(self):
        dialog = GenerateParametersDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            params = dialog.get_parameters()  # {'TB_NAME': ..., ...}

            # Сохраняем JSON
            try:
                GDocument.save_config_json(params)          # <-- теперь конфиг в JSON
                self._log(f"Параметры сохранены в config.json: {params}")
            except Exception as e:
                import traceback
                self._log(f"Ошибка сохранения config.json: {e}\n{traceback.format_exc()}")

            # Передаём параметры напрямую в генератор (лучший подход)
            # либо GDocument.generate сам пусть читает load_config_json() внутри.
            self.run_and_log(GDocument.create.create_all_templates)
            self.run_and_log(GDocument.generate.gen_all)

        else:
            self._log("Генерация отменена")

    def on_btn3_clicked(self):
        """Обработчик кнопки 3 — диалог выбора параметра"""
        dialog = BoolParameterDialog(self)

        if dialog.exec_() == QDialog.Accepted and getattr(dialog, "selected_value", None) is not None:
            self.run_and_log(GDocument.send_all, testing=dialog.selected_value)
        else:
            self._log("Выбор параметра отменен")


    def run_and_log(self, func, *args, **kwargs):
        """Выполняет функцию и пишет результат в историю"""
        try:
            result = func(*args, **kwargs)
            if result is None:
                self._log(f"Успешно: {func.__name__} выполнена")
            else:
                self._log(f"Успешно: {func.__name__} → {result}")
        except Exception as e:
            self._log(f"Ошибка в {func.__name__}: {e}\n{traceback.format_exc()}")

    def _log(self, msg: str):
        """Добавляет строку в историю"""
        self.history.append(msg)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.resize(400, 300)
    window.show()
    sys.exit(app.exec_())


