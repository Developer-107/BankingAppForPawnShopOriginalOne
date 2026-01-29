from utils import get_conn
from PyQt5.QtWidgets import (
    QWidget, QGridLayout, QLabel, QLineEdit, QPushButton,
    QMessageBox, QHBoxLayout
)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QSize

from utils import resource_path


class EditPaidPercentWindow(QWidget):
    def __init__(self, record_id):
        super().__init__()
        self.record_id = record_id
        self.setWindowTitle("პროცენტის გადახდის რედაქტირება")
        self.setWindowIcon(QIcon(resource_path("Icons/edit_data.png")))
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
        save_button.setIcon(QIcon(resource_path("Icons/save_icon.png")))
        save_button.setIconSize(QSize(35, 35))
        save_button.setStyleSheet("font-size: 16px;")
        save_button.clicked.connect(self.update_record)

        cancel_button = QPushButton("დახურვა")
        cancel_button.setIcon(QIcon(resource_path("Icons/cancel_icon.png")))
        cancel_button.setIconSize(QSize(35, 35))
        cancel_button.setStyleSheet("font-size: 16px;")
        cancel_button.clicked.connect(self.close)

        button_layout = QHBoxLayout()
        button_layout.addWidget(save_button)
        button_layout.addWidget(cancel_button)

        self.layout.addLayout(button_layout, len(fields), 1)
        self.setLayout(self.layout)

    def load_data(self):
        conn = get_conn()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM paid_percent_amount WHERE unique_id = %s", (self.record_id,))
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
            conn = get_conn()
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE paid_percent_amount SET
                    paid_amount = %s%s
                WHERE unique_id = %s%s
            """, (new_amount, self.record_id))

            cursor.execute("""
                            UPDATE paid_principle_and_paid_percentage_database SET
                                amount = %s%s
                            WHERE contract_id = %s%s AND status = %s%s AND date_of_inflow = %s
                        """, (new_amount, contract_id, status, payment_date))

            cursor.execute("""
                                        UPDATE inflow_order_only_percent_amount SET
                                            percent_paid_amount = %s, sum_of_money_paid= %s
                                        WHERE contract_id = %s AND payment_date = %s
                                    """, (new_amount, new_amount, contract_id, payment_date))

            cursor.execute("""
                                UPDATE inflow_order_both SET
                                    percent_paid_amount = %s
                                WHERE contract_id = %s AND payment_date = %s AND principle_paid_amount = %s
                                """, (new_amount, contract_id, payment_date, 0))
            conn.commit()
            conn.close()





            QMessageBox.information(self, "წარმატება", "ჩანაწერი განახლებულია")
            self.close()

        except Exception as e:
            QMessageBox.critical(self, "შეცდომა", f"შეცდომა განახლებაში:\n{e}")
