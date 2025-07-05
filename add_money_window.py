import sqlite3

from PyQt5.QtCore import QDate, QDateTime
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QWidget, QGridLayout, QLabel, QLineEdit, QDateEdit, QComboBox, QPushButton, QMessageBox


class AddMoney(QWidget):
    def __init__(self, contract_id, name_surname):
        super().__init__()
        self.contract_id = contract_id
        self.name_surname = name_surname
        self.setWindowTitle("თანხის დამატება")
        self.setWindowIcon(QIcon("Icons/add_money.png"))
        self.setFixedSize(400, 250)

        layout = QGridLayout()

        # Contract ID (readonly)
        layout.addWidget(QLabel("ხელშეკრულების №:"), 0, 0)
        self.contract_id_box = QLineEdit(str(contract_id))
        self.contract_id_box.setReadOnly(True)
        layout.addWidget(self.contract_id_box, 0, 1)

        # Name (readonly)
        layout.addWidget(QLabel("სახელი და გვარი: "), 1, 0)
        self.name_box = QLineEdit(str(name_surname))
        self.name_box.setReadOnly(True)
        layout.addWidget(self.name_box, 1, 1)

        # Add money input
        layout.addWidget(QLabel("დამატებული თანხა: "), 2, 0)
        self.added_money_amount = QLineEdit()
        layout.addWidget(self.added_money_amount, 2, 1)

        # Buttons
        save_button = QPushButton("შენახვა")
        save_button.clicked.connect(self.save_payment)
        cancel_button = QPushButton("დახურვა")
        cancel_button.clicked.connect(self.close)
        layout.addWidget(save_button, 3, 0)
        layout.addWidget(cancel_button, 3, 1)

        self.setLayout(layout)


    def save_payment(self):

        status_for_added_money = "დამატებული"
        date_of_addition = QDateTime.currentDateTime().toString("yyyy-MM-dd HH:mm:ss")
        contract_id = self.contract_id_box.text()

        try:
            # Step 1: Get the id_number from the original active_contracts table
            source_conn = sqlite3.connect("Databases/active_contracts.db")
            source_cursor = source_conn.cursor()
            source_cursor.execute("""SELECT id_number, additional_amounts, date, item_name, 
                                                 model, imei, given_money, tel_number, percent
                                            FROM active_contracts WHERE id = ?""", (contract_id,))
            result = source_cursor.fetchone()

            if not result:
                QMessageBox.warning(self, "შეცდომა", "მითითებული ID-ით ჩანაწერი ვერ მოიძებნა contracts ბაზაში.")
                return

            id_number_from_contracts = result[0]
            additional_amounts = result[1]
            contract_open_date = result[2]
            item_name = result[3]
            model = result[4]
            imei = result[5]
            given_money = result[6]
            tel_number = result[7]
            percent = result[8]

            updated_additional_amount = float(additional_amounts) + int(self.added_money_amount.text())
            new_added_percents = (given_money + updated_additional_amount) * percent / 100
            source_cursor.execute("""
                    UPDATE active_contracts
                    SET additional_amounts = ?, added_percents = ?
                    WHERE id = ?
                    """, (updated_additional_amount, new_added_percents,contract_id,))

            source_conn.commit()
            source_conn.close()



            conn = sqlite3.connect("Databases/given_and_additional_database.db")  # Make sure this matches your DB
            cursor = conn.cursor()


            # Insert in given_and_additional_database database
            cursor.execute("""
                INSERT INTO given_and_additional_database (
                    contract_id, date_of_outflow, name_surname, amount, status
                ) VALUES (?, ?, ?, ?, ?)
            """, (
                self.contract_id_box.text(),
                date_of_addition,
                self.name_box.text(),
                int(self.added_money_amount.text()),
                status_for_added_money
                ))

            conn.commit()
            conn.close()

            conn = sqlite3.connect("Databases/outflow_order.db")  # Make sure this matches your DB
            cursor = conn.cursor()

            # Insert in outflow_order database
            cursor.execute("""
                            INSERT INTO outflow_order (
                                contract_id, date, name_surname, tel_number, amount, status
                            ) VALUES (?, ?, ?, ?, ?, ?)
                        """, (
                self.contract_id_box.text(),
                date_of_addition,
                self.name_box.text(),
                tel_number,
                int(self.added_money_amount.text()),
                status_for_added_money
            ))

            conn.commit()
            conn.close()

            conn = sqlite3.connect("Databases/outflow_in_registry.db")  # Make sure this matches your DB
            cursor = conn.cursor()

            # Insert in outflow_in_registry database
            cursor.execute("""
                            INSERT INTO outflow_in_registry (
                                contract_id, date_of_C_O, name_surname, tel_number, id_number, item_name, model, IMEI,
                                given_money, date_of_addition, additional_amount, status
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (
                self.contract_id_box.text(),
                contract_open_date,
                self.name_box.text(),
                tel_number,
                id_number_from_contracts,
                item_name,
                model,
                imei,
                given_money,
                date_of_addition,
                int(self.added_money_amount.text()),
                status_for_added_money
            ))


            conn.commit()



            QMessageBox.information(self, "წარმატება", "მონაცემები შენახულია")
            self.close()
        except Exception as e:
            QMessageBox.critical(self, "შეცდომა", f"ვერ შევინახე მონაცემები:\n{e}")
        finally:
            conn.close()