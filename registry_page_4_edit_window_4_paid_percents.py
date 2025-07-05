import sqlite3
from PyQt5.QtWidgets import (
    QWidget, QGridLayout, QLabel, QLineEdit, QPushButton,
    QMessageBox, QHBoxLayout
)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QSize


class EditPaidPercentWindow(QWidget):
    def __init__(self, record_id):
        super().__init__()
        self.record_id = record_id
        self.setWindowTitle("პროცენტის გადახდის რედაქტირება")
        self.setWindowIcon(QIcon("Icons/edit_data.png"))
        self.resize(600, 300)
        self.build_ui()
        self.load_data()

    def build_ui(self):
        self.layout = QGridLayout()

        self.contract_id_box = QLineEdit()
        self.contract_id_box.setReadOnly(True)
        self.payment_date_box = QLineEdit()
        self.payment_date_box.setReadOnly(True)
        self.name_surname_box = QLineEdit()
        self.name_surname_box.setReadOnly(True)
        self.payment_amount_box = QLineEdit()
        self.status_box = QLineEdit()
        self.status_box.setReadOnly(True)

        fields = [
            ("კონტრაქტის ნომერი:", self.contract_id_box),
            ("გადახდის თარიღი:", self.payment_date_box),
            ("სახელი და გვარი:", self.name_surname_box),
            ("თანხა:", self.payment_amount_box),
            ("სტატუსი:", self.status_box)
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
        conn = sqlite3.connect("Databases/paid_percent_amount.db")
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM paid_percent_amount WHERE unique_id = ?", (self.record_id,))
        row = cursor.fetchone()
        conn.close()

        if row:
            self.contract_id_box.setText(str(row[1]))
            self.payment_date_box.setText(row[10])
            self.name_surname_box.setText(str(row[3]))
            self.payment_amount_box.setText(str(row[11]))
            self.status_box.setText(row[12])
            self.payment_amount_before = float(row[11])

    def update_record(self):
        try:
            new_amount = float(self.payment_amount_box.text())
            contract_id = self.contract_id_box.text()
            payment_date = self.payment_date_box.text()
            status = self.status_box.text()

            # 1. Update paid_percent_amount
            conn = sqlite3.connect("Databases/paid_percent_amount.db")
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE paid_percent_amount SET
                    paid_amount = ?
                WHERE unique_id = ?
            """, (new_amount, self.record_id))
            conn.commit()
            conn.close()

            conn = sqlite3.connect("Databases/paid_principle_and_paid_percentage_database.db")
            cursor = conn.cursor()
            cursor.execute("""
                            UPDATE paid_principle_and_paid_percentage_database SET
                                amount = ?
                            WHERE contract_id = ? AND status = ?
                        """, (new_amount, contract_id, status))
            conn.commit()
            conn.close()

            conn = sqlite3.connect("Databases/inflow_order_only_percent_amount.db")
            cursor = conn.cursor()
            cursor.execute("""
                                        UPDATE inflow_order_only_percent_amount SET
                                            percent_paid_amount = ?, sum_of_money_paid= ?
                                        WHERE contract_id = ? AND payment_date = ?
                                    """, (new_amount, new_amount, contract_id, payment_date))
            conn.commit()
            conn.close()

            conn = sqlite3.connect("Databases/inflow_order_both.db")
            cursor = conn.cursor()
            cursor.execute("""
                                UPDATE inflow_order_both SET
                                    percent_paid_amount = ?
                                WHERE contract_id = ? AND payment_date = ? AND principal_paid_money = ?
                                """, (new_amount, contract_id, payment_date, 0))
            conn.commit()
            conn.close()

            # U need to test this code in this GUI



            QMessageBox.information(self, "წარმატება", "ჩანაწერი განახლებულია")
            self.close()

        except Exception as e:
            QMessageBox.critical(self, "შეცდომა", f"შეცდომა განახლებაში:\n{e}")
