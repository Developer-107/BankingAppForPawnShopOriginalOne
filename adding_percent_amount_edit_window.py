import sqlite3
from PyQt5.QtWidgets import (
    QWidget, QGridLayout, QLabel, QLineEdit, QPushButton,
    QMessageBox, QHBoxLayout
)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QSize


class EditAddingPercentWindow(QWidget):
    def __init__(self, record_id):
        super().__init__()
        self.record_id = record_id
        self.setWindowTitle("დამატებული პროცენტის რედაქტირება")
        self.setWindowIcon(QIcon("Icons/edit_data.png"))
        self.resize(600, 300)
        self.build_ui()
        self.load_data()

    def build_ui(self):
        self.layout = QGridLayout()

        self.contract_id_box = QLineEdit()
        self.contract_id_box.setReadOnly(True)
        self.percent_addition_date_box = QLineEdit()
        self.percent_addition_date_box.setReadOnly(True)
        self.percent_amount_box = QLineEdit()
        self.name_surname_box = QLineEdit()
        self.name_surname_box.setReadOnly(True)
        self.id_number_box = QLineEdit()
        self.id_number_box.setReadOnly(True)

        fields = [
            ("ხელშეკრულების ID:", self.contract_id_box),
            ("თარიღი:", self.percent_addition_date_box),
            ("დამატებული პროცენტი:", self.percent_amount_box),
            ("სახელი და გვარი:", self.name_surname_box),
            ("პირადი ნომერი:", self.id_number_box)
        ]

        for i, (label_text, widget) in enumerate(fields):
            self.layout.addWidget(QLabel(label_text), i, 0)
            self.layout.addWidget(widget, i, 1)

        save_button = QPushButton("შენახვა")
        save_button.setIcon(QIcon("Icons/save_icon.png"))
        save_button.setIconSize(QSize(35, 35))
        save_button.setStyleSheet("font-size: 16px;")
        save_button.clicked.connect(self.update_record)

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
        conn = sqlite3.connect("Databases/adding_percent_amount.db")
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM adding_percent_amount WHERE unique_id = ?", (self.record_id,))
        row = cursor.fetchone()
        conn.close()

        if row:

            self.contract_id_box.setText(str(row[1]))
            self.percent_addition_date_box.setText(str(row[9]))
            self.percent_amount_box.setText(str(row[10]))
            self.name_surname_box.setText(str(row[3]))
            self.id_number_box.setText(str(row[5]))
            self.amount_before_editing = float(row[10])

    def update_record(self):
        try:
            new_amount = float(self.percent_amount_box.text())

            conn = sqlite3.connect("Databases/adding_percent_amount.db")
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE adding_percent_amount SET
                    percent_amount = ?
                WHERE contract_id = ? AND date_of_percent_addition = ?
            """, (
                new_amount,
                self.contract_id_box.text(),
                self.percent_addition_date_box.text()
            ))
            conn.commit()
            conn.close()

            conn = sqlite3.connect("Databases/active_contracts.db")
            cursor = conn.cursor()

            cursor.execute("SELECT added_percents FROM active_contracts WHERE id = ?", (self.contract_id_box.text(),))
            row = cursor.fetchone()
            conn.close()

            if row:
                added_percents_before = row[0]
            else:
                print("No record found")

            difference_between_last_and_new_amounts = -self.amount_before_editing + new_amount
            new_percents_amount = added_percents_before + difference_between_last_and_new_amounts

            conn = sqlite3.connect("Databases/active_contracts.db")
            cursor = conn.cursor()

            cursor.execute("""
                            UPDATE active_contracts SET
                                added_percents = ?
                            WHERE id = ?
                        """, (
                new_percents_amount,
                self.contract_id_box.text(),
            ))
            conn.commit()
            conn.close()


            QMessageBox.information(self, "წარმატება", "ჩანაწერი განახლდა წარმატებით")
            self.close()

        except Exception as e:
            QMessageBox.critical(self, "შეცდომა", f"შეცდომა განახლებაში:\n{e}")
