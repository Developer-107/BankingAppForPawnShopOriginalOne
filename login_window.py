from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLineEdit, QPushButton, QLabel, QMessageBox
from utils import get_conn

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

        login_button.setDefault(True)
        login_button.setAutoDefault(True)

        self.username_input.returnPressed.connect(self.check_login)
        self.password_input.returnPressed.connect(self.check_login)

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

        conn = get_conn()
        cursor = conn.cursor()
        cursor.execute("SELECT role, name_of_user, organisation, id_number_of_user FROM users WHERE username=%s AND password=%s", (username, password))
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
