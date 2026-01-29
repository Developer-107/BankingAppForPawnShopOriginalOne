from utils import get_conn

from PyQt5.QtCore import Qt, QDateTime
from PyQt5.QtWidgets import QWidget, QLabel, QLineEdit, QPushButton, QGridLayout, QMessageBox
from PyQt5.QtGui import QIcon

from utils import resource_path


class PaymentConfirmWindow(QWidget):
    def __init__(self, contract_id, name_surname, principal_paid, percent_paid, given_money, principal_should_be_paid):
        super().__init__()
        self.setWindowTitle("გადახდის დადასტურება")
        self.setWindowIcon(QIcon(resource_path("Icons/closed_contracts.png")))
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
        self.contract_id = contract_id
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
        self.remaining = principal_should_be_paid
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
        try:
            conn = get_conn()
            cursor = conn.cursor()

            cursor.execute("SELECT * FROM active_contracts WHERE id = %s", (self.contract_id,))
            row = cursor.fetchone()

            name_surname = str(row[3])
            id_number = str(row[4])
            tel_number = str(row[5])
            item_name = str(row[6])
            model = str(row[7])
            imei = str(row[8])
            percent = str(row[13])
            day_quantity = str(row[14])
            given_money = str(row[12])
            additional_amount = str(row[15])
            paid_principle = str(row[16])
            added_percents = str(row[18])
            paid_percents = str(row[19])
            status = str(row[20])
            comment = str(row[11])
            trusted_person = str(row[10])
            date_of_C_O = str(row[1])

            cursor.execute("""
                UPDATE active_contracts
                SET is_visible = 'დახურული'
                WHERE id = %s
            """, (self.contract_id,))

            # Insert in closed_contracts database
            cursor.execute("""
                            INSERT INTO closed_contracts (
                                id, contract_open_date, name_surname, id_number, tel_number, item_name, model, IMEI, percent,
                                percent_day_quantity, given_money, additional_money, paid_principle, added_percents,
                                paid_percents, status, date_of_closing, comment, trusted_person
                            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                        """, (
                self.contract_id,
                date_of_C_O,
                name_surname,
                id_number,
                tel_number,
                item_name,
                model,
                imei,
                percent,
                day_quantity,
                given_money,
                additional_amount,
                paid_principle,
                added_percents,
                paid_percents,
                status,
                QDateTime.currentDateTime().toString("yyyy-MM-dd HH:mm:ss"),
                comment,
                trusted_person
            ))

            conn.commit()
            conn.close()

            QMessageBox.information(self, "წარმატება", "გადახდა დადასტურებულია")
            self.close()

        except Exception as e:
            QMessageBox.critical(self, "შეცდომა", str(e))