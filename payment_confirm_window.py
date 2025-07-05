from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QWidget, QLabel, QLineEdit, QPushButton, QGridLayout, QMessageBox
from PyQt5.QtGui import QFont, QPalette, QColor, QIcon


class PaymentConfirmWindow(QWidget):
    def __init__(self, contract_id, name_surname, principal_paid, percent_paid, given_money):
        super().__init__()
        self.setWindowTitle("გადახდის დადასტურება")
        self.setWindowIcon(QIcon("Icons/closed_contracts.png"))
        self.setFixedSize(705, 250)

        layout = QGridLayout()

        # Header warning
        warning_label = QLabel("ყურადღება !!! თქვენ ნამდვილად გსურთ ამ ხელშეკრულების დახურვა?!")
        warning_label.setAlignment(Qt.AlignCenter)
        warning_label.setStyleSheet("font-weight: bold; font-size: 10pt; ")
        warning_label.setWordWrap(True)
        layout.addWidget(warning_label, 0, 0, 1, 4)

        # Labels and data
        layout.addWidget(QLabel(" ხელშეკრულების № "), 1, 0)
        contract_id_box = QLineEdit(str(contract_id))
        contract_id_box.setReadOnly(True)
        layout.addWidget(contract_id_box, 2, 0)

        layout.addWidget(QLabel(" სახელი და გვარი "), 1, 1)
        type_box = QLineEdit(str(name_surname))
        type_box.setReadOnly(True)
        layout.addWidget(type_box, 2, 1)

        layout.addWidget(QLabel(" გადახდილი ძირი თანხა "), 1, 2)
        paid_box = QLineEdit(str(principal_paid))
        paid_box.setReadOnly(True)
        layout.addWidget(paid_box, 2, 2)

        layout.addWidget(QLabel(" გადახდილი პროცენტი "), 1, 3)
        percent_box = QLineEdit(str(percent_paid))
        percent_box.setReadOnly(True)
        layout.addWidget(percent_box, 2, 3)

        # # Difference calculation
        self.remaining = float(given_money) - float(principal_paid)
        remaining_label = QLineEdit(str(self.remaining))
        remaining_label.setReadOnly(True)
        remaining_label.setAlignment(Qt.AlignCenter)
        remaining_label.setStyleSheet("""
            QLineEdit {
                background-color: #fffaf5;  /* light warm background for contrast */
                color: #4b0000;  /* very dark red, nearly black */
                font-size: 14px;
                font-weight: bold;
                border: 2px solid #4b0000;
                border-radius: 8px;
                padding: 3px;
            }
        """)
        layout.addWidget(QLabel("დარჩენილი გადასახდელი:"), 3, 0)
        layout.addWidget(remaining_label, 4, 0)

        # Buttons
        yes_button = QPushButton("დიახ")
        yes_button.clicked.connect(self.on_save_clicked)
        no_button = QPushButton("არა")
        no_button.clicked.connect(self.close)
        layout.addWidget(yes_button, 4, 2)
        layout.addWidget(no_button, 4, 3)

        self.setLayout(layout)

    def on_save_clicked(self):
        # Calculate remaining

        if self.remaining > 0:
            alert = QMessageBox(self)
            alert.setWindowTitle("ყურადღება")
            alert.setText(f"გთხოვთ დაფაროთ დარჩენილი თანხა : {self.remaining}")
            alert.setIcon(QMessageBox.Warning)
            alert.exec_()
            return  # Optionally stop further saving until full payment
        else:
            # Continue with save logic
            self.confirm_payment()

    def confirm_payment(self):
        # You can connect this to your DB update
        QMessageBox.information(self, "წარმატება", "გადახდა დადასტურებულია")
        self.close()