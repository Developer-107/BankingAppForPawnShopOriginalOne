from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLineEdit, QPushButton, QLabel, QMessageBox
import sqlite3

from utils import resource_path


class LoginWindow(QWidget):
    def __init__(self, app_callback):
        super().__init__()
        self.setWindowTitle("შესვლა")
        self.setWindowIcon(QIcon(resource_path("Icons/login_icon.png")))
        self.resize(400, 190)
        self.app_callback = app_callback



        self.username_input = QLineEdit()
        self.username_input.setPlaceholderText("მომხმარებელი")

        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_input.setPlaceholderText("პაროლი")

        login_button = QPushButton("შესვლა")
        login_button.clicked.connect(self.check_login)

        layout = QVBoxLayout()
        layout.addWidget(QLabel("მომხმარებელი"))
        layout.addWidget(self.username_input)
        layout.addWidget(QLabel("პაროლი"))
        layout.addWidget(self.password_input)
        layout.addWidget(login_button)
        self.setLayout(layout)

    def check_login(self):
        username = self.username_input.text()
        password = self.password_input.text()

        conn = sqlite3.connect(resource_path("Credentials/users.db"))
        cursor = conn.cursor()
        cursor.execute("SELECT role, name_of_user, organisation, id_number_of_user FROM users WHERE username=? AND password=?", (username, password))
        result = cursor.fetchone()
        conn.close()

        if result:
            role = result[0]
            name_of_user = result[1]
            organisation = result[2]
            id_number_of_user = result[3]
            self.app_callback(username, role, name_of_user, organisation, id_number_of_user)
            self.close()
        else:
            QMessageBox.warning(self, "Login Failed", "Invalid credentials.")
