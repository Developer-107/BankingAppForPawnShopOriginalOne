from utils import get_conn
from PyQt5.QtWidgets import (
    QWidget, QGridLayout, QLabel, QLineEdit, QPushButton,
    QMessageBox, QHBoxLayout
)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QSize

from utils import resource_path


class EditRegistryOutflowWindow(QWidget):
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

        # Fields based on inflows table
        self.date_box = QLineEdit()
        self.date_box.setReadOnly(True)
        self.amount_box = QLineEdit()
        self.name_surname_box = QLineEdit()
        self.name_surname_box.setReadOnly(True)
        self.id_number_box = QLineEdit()
        self.id_number_box.setReadOnly(True)
        self.status_box = QLineEdit()
        self.status_box.setReadOnly(True)

        fields = [
            ("თარიღი:", self.date_box),
            ("თანხა:", self.amount_box),
            ("სახელი და გვარი:", self.name_surname_box),
            ("პირადი ნომერი:", self.id_number_box),
            ("სტატუსი:", self.status_box)
        ]

        for i, (label_text, widget) in enumerate(fields):
            self.layout.addWidget(QLabel(label_text), i, 0)
            self.layout.addWidget(widget, i, 1)

        # Buttons
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
        conn = get_conn()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM outflow_in_registry WHERE unique_id = %s", (self.record_id,))
        row = cursor.fetchone()
        conn.close()

        if row:
            self.contract_id = row[1]
            self.date_box.setText(row[10])
            self.name_surname_box.setText(row[4])
            self.id_number_box.setText(row[5])
            self.amount_box.setText(str(row[11]))
            self.status_box.setText(str(row[12]))
            self.amount_before_editing = row[11]


    def update_record(self):
        try:
            conn = get_conn()
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE outflow_in_registry SET
                    date_of_addition = %s,
                    additional_amount = %s,
                    name_surname = %s,
                    id_number = %s,
                    status = %s
                WHERE unique_id = %s
            """, (
                self.date_box.text(),
                self.amount_box.text(),
                self.name_surname_box.text(),
                self.id_number_box.text(),
                self.status_box.text(),
                self.record_id
            ))

            cursor.execute("""
                            UPDATE given_and_additional_database SET
                                amount = %s
                            WHERE contract_id = %s and status = %s and amount = %s and date_of_outflow = %s
                        """, (
                self.amount_box.text(),
                self.contract_id,
                self.status_box.text(),
                self.amount_before_editing,
                self.date_box.text()
            ))

            cursor.execute("""
                                        UPDATE outflow_order SET
                                            amount = %s
                                        WHERE contract_id = %s and amount = %s and status = %s and date = %s
                                    """, (
                self.amount_box.text(),
                self.contract_id,
                self.amount_before_editing,
                self.status_box.text(),
                self.date_box.text()
            ))

            # Fetch old value
            cursor.execute("""
                SELECT additional_amounts, given_money, percent FROM active_contracts WHERE id = %s
            """, (self.contract_id,))
            row = cursor.fetchone()

            if row and row[0]:
                old_amount = float(row[0])
                given_money = float(row[1])
                percent = float(row[2])
            else:
                old_amount = 0.0  # Or whatever default makes sense

            new_amount = (-int(self.amount_before_editing) + float(self.amount_box.text()))
            new_additional_amounts = old_amount + new_amount
            new_added_percent = (given_money + new_additional_amounts) * percent / 100

            cursor.execute("""
                                                    UPDATE active_contracts SET
                                                        additional_amounts = %s,
                                                        added_percents = %s
                                                    WHERE id = %s
                                                """, (
                new_additional_amounts,
                new_added_percent,
                self.contract_id
            ))
            conn.commit()


            QMessageBox.information(self, "წარმატება", "ჩანაწერი განახლებულია")
            self.close()
        except Exception as e:
            QMessageBox.critical(self, "შეცდომა", f"შეცდომა განახლებაში:\n{e}")
        finally:
            conn.close()