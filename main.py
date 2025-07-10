import sqlite3
import sys

from PyQt5.QtCore import QSize, Qt, QDateTime
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QGridLayout, QWidget, QLabel, QGroupBox, \
    QVBoxLayout, QToolButton

from active_contracts_window import ActiveContracts
from clients_in_the_black_list_window import ClientsInTheBlackList
from closed_contracts_window import ClosedContracts
from contract_registry_window import ContractRegistry
from help_window import HelpWindow
from login_window import LoginWindow
from money_control_window import MoneyControl


# Subclass QMainWindow to customize your application's main window
class MainWindow(QMainWindow):
    def __init__(self, username, role, name_of_user, organisation, id_number_of_user):
        super().__init__()
        self.username = username
        self.role = role
        self.name_of_user = name_of_user
        self.organisation = organisation
        self.id_number_of_user = id_number_of_user

        self.money_control_window = None
        self.clients_in_the_black_list_window = None
        self.closed_contracts_window = None
        self.contract_registry_window = None
        self.active_contracts_window = None
        self.help_window = None
        self.setWindowTitle("ლომბარდი")
        self.setWindowIcon(QIcon("Icons/app_icon.png"))

        layout = QGridLayout()
        layout.setContentsMargins(10, 40, 10, 140)

        # Buttons for other pages
        # Active contracts settings
        active_contracts = QToolButton()
        active_contracts.setText("\n    მოქმედი ხელშეკრულებები    ")
        active_contracts.setIcon(QIcon("Icons/contract_icon.png"))
        active_contracts.setIconSize(QSize(80, 80))
        active_contracts.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        active_contracts.setStyleSheet("font-size: 16px; font-weight: bold;")
        active_contracts.clicked.connect(self.open_active_contracts_window)

        # Contract registry settings
        contract_registry = QToolButton()
        contract_registry.setText("\n  ხელშეკრულებების რეესტრი   ")
        contract_registry.setIcon(QIcon("Icons/contract_registry.png"))
        contract_registry.setIconSize(QSize(80, 80))
        contract_registry.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        contract_registry.setStyleSheet("font-size: 16px; font-weight: bold;")
        contract_registry.clicked.connect(self.open_contract_registry_window)

        # Closed contracts settings
        closed_contracts = QToolButton()
        closed_contracts.setText("\n  დახურული ხელშეკრულებები   ")
        closed_contracts.setIcon(QIcon("Icons/closed_contracts.png"))
        closed_contracts.setIconSize(QSize(80, 80))
        closed_contracts.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        closed_contracts.setStyleSheet("font-size: 16px; font-weight: bold;")
        closed_contracts.clicked.connect(self.open_closed_contracts_window)

        # clients in black list settings
        clients_in_the_black_list = QToolButton()
        clients_in_the_black_list.setText("\n შავ სიაში მყოფი კლიენტები  ")
        clients_in_the_black_list.setIcon(QIcon("Icons/blacklist.png"))
        clients_in_the_black_list.setIconSize(QSize(80, 80))
        clients_in_the_black_list.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        clients_in_the_black_list.setStyleSheet("font-size: 16px; font-weight: bold;")
        clients_in_the_black_list.clicked.connect(self.open_clients_in_the_black_list_window)

        # Rate the client settings
        money_control = QToolButton()
        money_control.setText("\n   თანხების კონტროლი    ")
        money_control.setIcon(QIcon("Icons/money_control.png"))
        money_control.setIconSize(QSize(80, 80))
        money_control.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        money_control.setStyleSheet("font-size: 16px; font-weight: bold;")
        money_control.clicked.connect(self.open_money_control_window)

        # Help button settings
        help_button = QToolButton()
        help_button.setText("\n         დახმარება         ")
        help_button.setIcon(QIcon("Icons/help.png"))
        help_button.setIconSize(QSize(80, 80))
        help_button.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        help_button.setStyleSheet("font-size: 16px; font-weight: bold;")
        help_button.clicked.connect(self.open_help_window)


        # Button layout
        layout.addWidget(active_contracts, 0, 0)
        layout.addWidget(contract_registry, 0, 1)
        layout.addWidget(closed_contracts, 0, 2)
        layout.addWidget(money_control, 0, 3)
        layout.addWidget(clients_in_the_black_list, 0, 4)
        layout.addWidget(help_button, 0, 5)


        # Box
        box = QGroupBox("პროგრამაში შესულია")
        box.setStyleSheet("QGroupBox {font-style: italic; font-size: 10pt; }")
        box_layout = QGridLayout()


        # Labels
        logged_in_the_system = QLabel(" ოპერატორი:")
        font = logged_in_the_system.font()
        font.setPointSize(14)
        logged_in_the_system.setFont(font)
        box_layout.addWidget(logged_in_the_system, 0, 0)

        logged_in_the_system_box = QLabel("")
        logged_in_the_system_box.setText(self.name_of_user)
        logged_in_the_system_box.setStyleSheet("border: 1px solid gray; padding: 5px;")
        logged_in_the_system_box.setFixedHeight(35)
        # logged_in_the_system_box.setFixedWidth(140)
        box_layout.addWidget(logged_in_the_system_box, 0, 1, 1, 2)

        operational_access = QLabel(" მინიჭებული უფლება:")
        font1 = operational_access.font()
        font1.setPointSize(14)
        operational_access.setFont(font)
        box_layout.addWidget(operational_access, 1, 0)

        operational_access_box = QLabel("")
        operational_access_box.setText(self.role)
        operational_access_box.setStyleSheet("border: 1px solid gray; padding: 5px;")
        operational_access_box.setFixedHeight(35)
        # operational_access_box.setFixedWidth(140)
        box_layout.addWidget(operational_access_box, 1, 1)

        operational_access_box1 = QLabel("")
        operational_access_box1.setText(QDateTime.currentDateTime().toString("yyyy-MM-dd HH:mm:ss"))
        operational_access_box1.setStyleSheet("border: 1px solid gray; padding: 5px;")
        operational_access_box1.setFixedHeight(35)
        operational_access_box1.setFixedWidth(190)
        box_layout.addWidget(operational_access_box1, 1, 2)


        organisation_name_label = QLabel(" ორგანიზაციის სახელი:")
        font2 = organisation_name_label.font()
        font2.setPointSize(14)
        organisation_name_label.setFont(font)
        box_layout.addWidget(organisation_name_label, 2, 0)

        organisation_name_label = QLabel("")
        organisation_name_label.setText(self.organisation)
        organisation_name_label.setStyleSheet("border: 1px solid gray; padding: 5px;")
        organisation_name_label.setFixedHeight(35)
        # organisation_name_label.setFixedWidth(140)
        box_layout.addWidget(organisation_name_label, 2, 1, 1, 2)



        # Set layout for the box and add to the main layout
        box.setLayout(box_layout)
        empty_widget = QWidget()
        empty_widget.setFixedHeight(80)  # height of the "br" space

        layout.addWidget(empty_widget, 1, 1, 1, 1)  # blank row spanning columns

        layout.addWidget(box, 2, 1, 1, 4)


        # Window size
        self.resize(1700, 800)

        # Set the central widget of the Window.
        widget = QWidget()
        widget.setLayout(layout)
        self.setCentralWidget(widget)


    def open_help_window(self):
        self.help_window = HelpWindow()
        self.help_window.show()

    def open_active_contracts_window(self):
        self.active_contracts_window = ActiveContracts(self.role, self.name_of_user,
                                                       self.organisation,
                                                       self.id_number_of_user)
        self.active_contracts_window.show()

    def open_contract_registry_window(self):
        self.contract_registry_window = ContractRegistry(self.role, self.name_of_user, self.organisation)
        self.contract_registry_window.show()

    def open_closed_contracts_window(self):
        self.closed_contracts_window = ClosedContracts(self.role)
        self.closed_contracts_window.show()

    def open_clients_in_the_black_list_window(self):
        self.clients_in_the_black_list_window = ClientsInTheBlackList(self.role)
        self.clients_in_the_black_list_window.show()

    def open_money_control_window(self):
        self.money_control_window = MoneyControl()
        self.money_control_window.show()


conn = sqlite3.connect("Credentials/users.db")
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT,
    password TEXT,
    name_of_user TEXT,
    organisation TEXT,
    id_number_of_user TEXT,
    role TEXT CHECK(role IN ('admin', 'user'))
)
""")


def launch_main(username, role, name_of_user, organisation, id_number_of_user):
    global main_window  # prevent it from being garbage collected
    main_window = MainWindow(username, role, name_of_user, organisation, id_number_of_user)
    main_window.show()



app = QApplication(sys.argv)
login = LoginWindow(app_callback=launch_main)
login.show()
sys.exit(app.exec_())