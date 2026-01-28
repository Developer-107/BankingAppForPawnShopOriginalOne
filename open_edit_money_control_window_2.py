from utils import get_conn
from PyQt5.QtWidgets import (
    QWidget, QGridLayout, QLabel, QLineEdit, QPushButton,
    QMessageBox, QHBoxLayout
)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QSize

from utils import resource_path


class EditInPrincipalInflowsInRegistryWindow(QWidget):
    def __init__(self, record_id):
        super().__init__()
        self.record_id = record_id
        self.setWindowTitle("შემოსვლის ჩანაწერის რედაქტირება")
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
        cursor.execute("SELECT * FROM paid_principle_registry WHERE unique_id = %s", (self.record_id,))
        row = cursor.fetchone()
        conn.close()

        if row:
            self.payment_date_box.setText(row[10])
            self.payment_amount_box.setText(str(row[11]))
            self.name_surname_box.setText(str(row[3]))
            self.contract_id_box.setText(str(row[1]))
            self.status_box.setText(str(row[12]))
            self.payment_amount_before = row[11]

    def update_record(self):
        try:
            conn = get_conn()
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE paid_principle_registry SET
                    payment_amount = %s
                WHERE unique_id = %s
            """, (
                self.payment_amount_box.text(),
                self.record_id
            ))
            conn.commit()
            conn.close()

            conn = get_conn()
            cursor = conn.cursor()
            cursor.execute("""
                            UPDATE paid_principle_and_paid_percentage_database SET
                                amount = %s
                            WHERE contract_id = %s and date_of_inflow = %s and status = %s
                        """, (
                self.payment_amount_box.text(),
                self.contract_id_box.text(),
                self.payment_date_box.text(),
                self.status_box.text()
            ))
            conn.commit()
            conn.close()

            conn = get_conn()
            cursor = conn.cursor()
            cursor.execute("""
                                UPDATE inflow_order_only_principal_amount SET
                                principle_paid_amount = %s, sum_of_money_paid = %s
                                WHERE contract_id = %s and payment_date = %s
                           """, (
                self.payment_amount_box.text(),
                self.payment_amount_box.text(),
                self.contract_id_box.text(),
                self.payment_date_box.text()
            ))
            conn.commit()
            conn.close()

            conn = get_conn()
            cursor = conn.cursor()
            cursor.execute("""
                               UPDATE inflow_order_both SET
                               principle_paid_amount = %s
                               WHERE contract_id = %s and payment_date = %s and percent_paid_amount = %s
                           """, (
                self.payment_amount_box.text(),
                self.contract_id_box.text(),
                self.payment_date_box.text(),
                0
            ))
            conn.commit()
            conn.close()

            conn = get_conn()
            cursor = conn.cursor()
            cursor.execute("""
                SELECT principal_paid FROM active_contracts WHERE id = %s
            """, (self.contract_id_box.text(),))
            row = cursor.fetchone()
            conn.close()

            if row:
                principal_paid_before = row[0]



            difference_between_changed_and_before = -float(self.payment_amount_before) + float(self.payment_amount_box.text())
            principal_paid_new_amount = principal_paid_before + difference_between_changed_and_before


            conn = get_conn()
            cursor = conn.cursor()
            cursor.execute("""
                            UPDATE active_contracts SET
                                principal_paid = %s
                            WHERE id = %s
                        """, (
                principal_paid_new_amount,
                self.contract_id_box.text(),
            ))
            conn.commit()
            conn.close()



            QMessageBox.information(self, "წარმატება", "ჩანაწერი განახლებულია")
            self.close()
        except Exception as e:
            QMessageBox.critical(self, "შეცდომა", f"შეცდომა განახლებაში:\n{e}")


