import GDocument
import config

import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QTextEdit

# === Функции с параметрами ===
def fun1(x, y):
    return f"fun1: {x} + {y} = {x + y}"

def fun2(name):
    return f"fun2: Привет, {name}!"

def fun3(a, b, c=0):
    return f"fun3: {a} * {b} + {c} = {a * b + c}"


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Функции с параметрами и история")
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Поле истории
        self.history = QTextEdit()
        self.history.setReadOnly(True)
        self.history.setPlaceholderText("История выполнений появится здесь...")
        layout.addWidget(self.history)

        # Кнопка 1 — передаем параметры 3 и 5
        btn1 = QPushButton("Создать шаблоны")
        btn1.clicked.connect(lambda: self.run_and_log(GDocument.create.create_all_templates))
        layout.addWidget(btn1)

        # Кнопка 2 — передаем имя
        btn2 = QPushButton("Сгенерировать ")
        btn2.clicked.connect(lambda: self.run_and_log(GDocument.generate.gen_all, config.FILES_FOLDER_NAME))
        layout.addWidget(btn2)

        # Кнопка 3 — передаем несколько параметров
        btn3 = QPushButton("Кнопка3")
        btn3.clicked.connect(lambda: self.run_and_log(fun3, 2, 4, c=7))
        layout.addWidget(btn3)

        self.setLayout(layout)

    def run_and_log(self, func, *args, **kwargs):
        """Выполняет функцию с аргументами и пишет результат в историю."""
        try:
            result = func(*args, **kwargs)
            self.history.append(str(result))
        except Exception as e:
            self.history.append(f"Ошибка: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.resize(400, 300)
    window.show()
    sys.exit(app.exec_())


