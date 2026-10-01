import GDocument
import sys
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QApplication, QMessageBox, QLineEdit, QWidget, QVBoxLayout, QPushButton, QTextEdit, QDialog, QLabel, QHBoxLayout, QDialogButtonBox, QFormLayout
import traceback
import config

class GenerateParametersDialog(QDialog):
    def __init__(self, parent=None, defaults=None):
        super().__init__(parent)
        self.setWindowTitle("Информация о мероприятии")
        self.setModal(True)
        self.defaults = defaults or {}
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        def row(label_text, key):
            h = QHBoxLayout()
            h.addWidget(QLabel(label_text))
            edit = QLineEdit()
            edit.setPlaceholderText(key)
            edit.setText(str(self.defaults.get(key, "")))  # последние введенные значения
            h.addWidget(edit)
            layout.addLayout(h)
            return edit

        self.event_name_edit= row("Краткое название мероприятия (н-р, NPW-2025):", "EVENT_NAME")
        self.event_info_edit= row("Полное название мероприятия (И.п., «что?», н-р, Научная школа):", "EVENT_INFO")
        self.event_prep_edit= row("  в предложном падеже («в чём?»):", "EVENT_INFO_PREP")
        self.event_gen_edit = row("  в родительном падеже («чего?»):", "EVENT_INFO_GEN")
        self.date_info_edit = row("Даты проведения (н-р, 7–13 сентября):", "DATE_INFO")
        self.place_info_edit= row("Место проведения (н-р, г. Москва):", "PLACE_INFO")
        self.oferta_link_edit=row("Ссылка на Публичную оферту:", "OFERTA_LINK")

        # Старый config.json: в EVENT_INFO лежало название уже в предложном падеже
        if self.defaults.get('EVENT_INFO') and not self.defaults.get('EVENT_INFO_PREP'):
            self.event_prep_edit.setText(self.defaults['EVENT_INFO'])
            self.event_info_edit.clear()

        # Падежи пересчитываются при вводе полного названия, их можно поправить вручную
        self.event_info_edit.textEdited.connect(self.update_cases)

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

    def update_cases(self, text):
        try:
            cases = GDocument.event_cases(text.strip())
        except Exception as e:  # словари pymorphy3 не найдены и т.п. — вводим вручную
            self.event_prep_edit.setPlaceholderText(f"Не удалось просклонять: {e}")
            return
        self.event_prep_edit.setText(cases['EVENT_INFO_PREP'])
        self.event_gen_edit.setText(cases['EVENT_INFO_GEN'])

    def on_ok(self):
        # Все поля попадают в шаблоны, поэтому пустые не допускаем
        if any(not v for v in self.get_parameters().values()):
            QMessageBox.warning(self, "Проверка", "Заполните все поля.")
            return
        self.accept()

    def get_parameters(self):
        """Возвращает параметры в виде словаря с нужными ключами."""
        return {
            'EVENT_NAME': self.event_name_edit.text().strip(),
            'EVENT_INFO': self.event_info_edit.text().strip(),
            'EVENT_INFO_PREP': self.event_prep_edit.text().strip(),
            'EVENT_INFO_GEN': self.event_gen_edit.text().strip(),
            'DATE_INFO': self.date_info_edit.text().strip(),
            'PLACE_INFO': self.place_info_edit.text().strip(),
            'OFERTA_LINK': self.oferta_link_edit.text().strip(),
        }

class LoginDialog(QDialog):
    """Запрос логина и пароля почты, с которой идет рассылка."""
    def __init__(self, parent=None, login=""):
        super().__init__(parent)
        self.setWindowTitle("Вход в почту для рассылки")
        self.setModal(True)

        layout = QVBoxLayout()
        form = QFormLayout()

        self.login_edit = QLineEdit(login)
        self.login_edit.setPlaceholderText("name@mail.ru")
        form.addRow("Логин (email):", self.login_edit)

        self.password_edit = QLineEdit()
        self.password_edit.setEchoMode(QLineEdit.Password)  # пароль скрыт точками
        form.addRow("Пароль:", self.password_edit)
        layout.addLayout(form)

        hint = QLabel("Для mail.ru нужен пароль для внешних приложений.\nПароль не сохраняется на диске.")
        hint.setStyleSheet("color: gray")
        layout.addWidget(hint)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Cancel).setText("Отмена")
        buttons.accepted.connect(self.on_ok)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.setLayout(layout)

        # курсор сразу в пустое поле
        (self.password_edit if login else self.login_edit).setFocus()

    def on_ok(self):
        if not self.login_edit.text().strip() or not self.password_edit.text():
            QMessageBox.warning(self, "Проверка", "Введите логин и пароль.")
            return
        self.accept()

    def get_credentials(self):
        return self.login_edit.text().strip(), self.password_edit.text()

class BoolParameterDialog(QDialog):
    def __init__(self, parent=None, test_address=""):
        super().__init__(parent)
        self.test_address = test_address
        self.setWindowTitle("Выбор параметров отправки")
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

        test_info_layout = QHBoxLayout()
        # Подпись под кнопкой
        self.test_label = QLabel(f"Тестовое письмо будет отправлено на: {self.test_address}")
        test_info_layout.addWidget(self.test_label)

        layout.addLayout(button_layout)
        layout.addLayout(test_info_layout)
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
        answer = QMessageBox.question(
            self, "Подтверждение",
            "Отправить письма ВСЕМ участникам из таблицы?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if answer != QMessageBox.Yes:
            return
        self.selected_value = False
        self.accept()

class Worker(QThread):
    """Выполняет долгую функцию в фоне, чтобы окно не «зависало»."""
    done = pyqtSignal(str)

    def __init__(self, func, *args, **kwargs):
        super().__init__()
        self.func, self.args, self.kwargs = func, args, kwargs

    def run(self):
        # docx2pdf работает с Word через COM, а COM нужно инициализировать в каждом потоке
        pythoncom = None
        if sys.platform.startswith("win"):
            try:
                import pythoncom
                pythoncom.CoInitialize()
            except ImportError:
                pythoncom = None
        try:
            result = self.func(*self.args, **self.kwargs)
            name = self.func.__name__
            if result is None:
                self.done.emit(f"Успешно: {name} выполнена")
            else:
                self.done.emit(f"Успешно: {name} → {result}")
        except Exception as e:
            self.done.emit(f"Ошибка в {self.func.__name__}: {e}\n{traceback.format_exc()}")
        finally:
            if pythoncom is not None:
                pythoncom.CoUninitialize()

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("DocApp")
        self.worker = None
        # учетные данные почты: хранятся только в памяти, пока открыта программа
        self.mail_login = config.FROM_MAIL
        self.mail_password = ""
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        self.history = QTextEdit()
        self.history.setReadOnly(True)
        self.history.setPlaceholderText("История выполнений появится здесь...")
        layout.addWidget(self.history)

        btn1 = QPushButton("Создать шаблоны")
        btn1.clicked.connect(self.on_btn1_clicked)
        layout.addWidget(btn1)

        btn2 = QPushButton("Сгенерировать")
        btn2.clicked.connect(lambda: self.run_and_log(GDocument.generate.gen_all))
        layout.addWidget(btn2)

        btn3 = QPushButton("Рассылка")
        btn3.clicked.connect(self.on_btn3_clicked)
        layout.addWidget(btn3)

        self.buttons = [btn1, btn2, btn3]
        self.setLayout(layout)

    def on_btn1_clicked(self):
        try:
            defaults = GDocument.load_config()
        except Exception as e:
            self._log(f"Не удалось прочитать config.json: {e}")
            defaults = {}

        dialog = GenerateParametersDialog(self, defaults)
        if dialog.exec_() != QDialog.Accepted:
            self._log("Создание шаблонов отменено")
            return

        params = dialog.get_parameters()

        # Сохраняем JSON
        try:
            GDocument.save_config(params)
            self._log(f"Параметры сохранены: {params}")
        except Exception as e:
            self._log(f"Ошибка сохранения config.json: {e}\n{traceback.format_exc()}")
            return

        self.run_and_log(GDocument.create.create_all_templates)

    def on_btn3_clicked(self):
        """Обработчик кнопки 3 — вход в почту и выбор типа рассылки"""
        login_dialog = LoginDialog(self, self.mail_login)
        if login_dialog.exec_() != QDialog.Accepted:
            self._log("Рассылка отменена")
            return
        self.mail_login, self.mail_password = login_dialog.get_credentials()

        dialog = BoolParameterDialog(self, test_address=self.mail_login)
        if dialog.exec_() == QDialog.Accepted and dialog.selected_value is not None:
            self.run_and_log(GDocument.send_all, testing=dialog.selected_value,
                             login=self.mail_login, password=self.mail_password)
        else:
            self._log("Выбор параметра отменен")

    def run_and_log(self, func, *args, **kwargs):
        """Запускает функцию в фоне и пишет результат в историю"""
        if self.worker is not None and self.worker.isRunning():
            self._log("Дождитесь окончания текущей операции")
            return

        self._set_busy(True)
        self._log(f"Запущено: {func.__name__} ...")
        self.worker = Worker(func, *args, **kwargs)
        self.worker.done.connect(self._log)
        self.worker.finished.connect(lambda: self._set_busy(False))
        self.worker.start()

    def _set_busy(self, busy: bool):
        for b in self.buttons:
            b.setEnabled(not busy)

    def _log(self, msg: str):
        """Добавляет строку в историю"""
        self.history.append(msg)

    def closeEvent(self, event):
        if self.worker is not None and self.worker.isRunning():
            QMessageBox.warning(self, "DocApp", "Дождитесь окончания текущей операции.")
            event.ignore()
            return
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.resize(400, 300)
    window.show()
    sys.exit(app.exec_())
