import sqlite3
from PyQt5.QtWidgets import (
    QWidget, QGridLayout, QLabel, QLineEdit, QPushButton,
    QMessageBox, QHBoxLayout, QComboBox
)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QSize, Qt, QDateTime


class EditWindow(QWidget):
    def __init__(self, record_id, role):
        super().__init__()
        self.record_id = record_id
        self.setWindowTitle("რედაქტირება")
        self.setWindowIcon(QIcon("Icons/edit_data.png"))
        self.resize(1400, 500)
        self.role = role
        self.build_ui()
        self.load_data()

    def build_ui(self):
        self.layout = QGridLayout()

        # Labels and Fields
        self.name_surname_box = QLineEdit()
        self.id_number_box = QLineEdit()
        self.tel_number_box = QLineEdit()
        self.item_name_box = QLineEdit()
        self.model_box = QLineEdit()
        self.imei_sn_box = QLineEdit()
        self.choose_the_type_box = QLineEdit()
        self.choose_the_type_box.setEnabled(False)
        self.trusted_person_box = QLineEdit()
        self.comment_box = QLineEdit()
        self.given_money_box = QLineEdit()
        if self.role != "admin":
            self.given_money_box.setEnabled(False)
        # --- Percent ComboBox ---
        self.percent_box = QComboBox()
        self.percent_box.addItems(["2.5", "5", "10", "15"])
        self.percent_box.setCurrentText("10")
        self.percent_box.setStyleSheet("padding: 5px; font-size: 14px;")
        if self.role != "admin":
            self.percent_box.setEnabled(False)

        # --- Day Quantity ComboBox ---
        self.day_quantity_box = QComboBox()
        self.day_quantity_box.addItems(["10", "15", "30"])
        self.day_quantity_box.setCurrentText("10")
        self.day_quantity_box.setStyleSheet("padding: 5px; font-size: 14px;")
        if self.role != "admin":
            self.day_quantity_box.setEnabled(False)



        fields = [
            ("სახელი და გვარი:", self.name_surname_box),
            ("პირადი ნომერი:", self.id_number_box),
            ("ტელეფონი:", self.tel_number_box),
            ("ნივთი:", self.item_name_box),
            ("მოდელი:", self.model_box),
            ("IMEI:", self.imei_sn_box),
            ("ტიპი:", self.choose_the_type_box),
            ("მინდობილი პირი:", self.trusted_person_box),
            ("კომენტარი:", self.comment_box),
            ("გაცემული თანხა:", self.given_money_box),
            ("პროცენტი:", self.percent_box),
            ("დღეების რაოდენობა:", self.day_quantity_box),
        ]

        for i, (label_text, widget) in enumerate(fields):
            self.layout.addWidget(QLabel(label_text), i, 0)
            self.layout.addWidget(widget, i, 1)

        # Save Button
        save_button = QPushButton("შენახვა")
        save_button.setIcon(QIcon("Icons/save_icon.png"))
        save_button.setIconSize(QSize(35, 35))
        save_button.setStyleSheet("font-size: 16px;")
        save_button.clicked.connect(self.update_record)

        # Cancel Button
        cancel_button = QPushButton("დახურვა")
        cancel_button.setIcon(QIcon("Icons/cancel_icon.png"))
        cancel_button.setIconSize(QSize(35, 35))
        cancel_button.setStyleSheet("font-size: 16px;")
        cancel_button.clicked.connect(self.close)

        button_layout = QHBoxLayout()
        button_layout.addWidget(save_button)
        button_layout.addWidget(cancel_button)

        self.layout.addLayout(button_layout, len(fields), 1)
        self.setLayout(self.layout)

    def load_data(self):
        conn = sqlite3.connect("Databases/active_contracts.db")
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM active_contracts_view WHERE id = ?", (self.record_id,))
        row = cursor.fetchone()
        conn.close()

        if row:
            self.name_surname_box.setText(row[3])
            self.id_number_box.setText(row[4])
            self.tel_number_box.setText(row[5])
            self.item_name_box.setText(row[6])
            self.model_box.setText(row[7])
            self.imei_sn_box.setText(row[8])
            self.choose_the_type_box.setText(row[9])
            self.trusted_person_box.setText(row[10])
            self.comment_box.setText(row[11])
            self.given_money_box.setText(str(int(row[12])))
            self.percent_box.setCurrentText(str(row[13]))  # ✅
            self.day_quantity_box.setCurrentText(str(row[14]))  # ✅
            self.additional_amounts = row[15]
            self.given_money_before = row[11]
            self.date_of_outflow = row[1]


    def update_record(self):
        try:
            # 1. Update active_contracts.db
            conn = sqlite3.connect("Databases/active_contracts.db")
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE active_contracts SET
                    name_surname = ?,
                    id_number = ?,
                    tel_number = ?,
                    item_name = ?,
                    model = ?,
                    imei = ?,
                    type = ?,
                    trusted_person = ?,
                    comment = ?,
                    given_money = ?,
                    percent = ?,
                    day_quantity = ?,
                    added_percents = ?
                WHERE id = ?
            """, (
                self.name_surname_box.text(),
                self.id_number_box.text(),
                self.tel_number_box.text(),
                self.item_name_box.text(),
                self.model_box.text(),
                self.imei_sn_box.text(),
                self.choose_the_type_box.text(),
                self.trusted_person_box.text(),
                self.comment_box.text(),
                float(self.given_money_box.text()),
                float(self.percent_box.currentText()),
                int(self.day_quantity_box.currentText()),
                float((int(self.given_money_box.text()) + self.additional_amounts) * (float(self.percent_box.currentText()) / 100)),
                self.record_id
            ))
            conn.commit()
            conn.close()

            status = 'გაცემული ძირი თანხა'

            # 2. Update given_and_additional_database.db
            conn = sqlite3.connect("Databases/given_and_additional_database.db")
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE given_and_additional_database SET
                    name_surname = ?,
                    amount = ?
                WHERE contract_id = ? AND status = ? and date_of_outflow = ?
            """, (
                self.name_surname_box.text(),
                self.given_money_box.text(),
                self.record_id,
                status,
                self.date_of_outflow
            ))
            conn.commit()
            conn.close()


            office_mob_number = "599 222 918"

            conn = sqlite3.connect("Databases/contracts.db")
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE contracts SET
                    name_surname = ?,
                    id_number = ?,
                    tel_number = ?,
                    item_name = ?,
                    model = ?,
                    IMEI = ?,
                    given_money = ?,
                    percent_day_quantity = ?,
                    first_added_percent = ?,
                    office_mob_number = ?
                WHERE contract_id = ?
            """, (
                self.name_surname_box.text(),
                self.id_number_box.text(),
                self.tel_number_box.text(),
                self.item_name_box.text(),
                self.model_box.text(),
                self.imei_sn_box.text(),
                self.given_money_box.text(),
                int(self.day_quantity_box.currentText()),
                float((int(self.given_money_box.text()) + self.additional_amounts) * (float(self.percent_box.currentText()) / 100)),
                office_mob_number,
                self.record_id
            ))
            conn.commit()
            conn.close()

            # 4. Update outflow_order.db
            conn = sqlite3.connect("Databases/outflow_order.db")
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE outflow_order SET
                    name_surname = ?,
                    tel_number = ?,
                    amount = ?
                WHERE contract_id = ? and amount = ? and status = ? and date = ?
            """, (
                self.name_surname_box.text(),
                self.tel_number_box.text(),
                self.given_money_box.text(),
                self.record_id,
                self.given_money_before,
                status,
                self.date_of_outflow
            ))
            conn.commit()
            conn.close()


            # 5. Update adding_percent_amount.db
            conn = sqlite3.connect("Databases/adding_percent_amount.db")
            cursor = conn.cursor()
            cursor.execute("""
                           UPDATE adding_percent_amount SET
                               percent_amount = ?
                           WHERE contract_id = ? AND date_of_percent_addition = ?
                       """, (
                float((int(self.given_money_box.text()) + self.additional_amounts) * (float(self.percent_box.currentText()) / 100)),
                self.record_id,
                self.date_of_outflow
            ))
            conn.commit()
            conn.close()




            QMessageBox.information(self, "წარმატება", "ჩანაწერები განახლებულია")
            self.close()

        except Exception as e:
            QMessageBox.critical(self, "შეცდომა", f"შეცდომა განახლებაში:\n{e}")

        finally:
            try:
                conn.close()
            except:
                pass
