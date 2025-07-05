import sqlite3
import sys
from datetime import datetime, timedelta

from PyQt5.QtCore import Qt, QSize, QDate
from PyQt5.QtGui import QIcon
from PyQt5.QtSql import QSqlDatabase, QSqlTableModel
from PyQt5.QtWidgets import QWidget, QGridLayout, QGroupBox, QToolButton, QCheckBox, QLabel, QRadioButton, QButtonGroup, \
    QLineEdit, QDateEdit, QPushButton, QTableView, QAbstractItemView, QMessageBox, QMenu, QAction

import tempfile
import pandas as pd
import os
import subprocess
from PyQt5.QtWidgets import QMessageBox

from add_money_window import AddMoney
from open_edit_window import EditWindow
from open_add_window import AddWindow
from payment_confirm_window import PaymentConfirmWindow
from payment_window import PaymentWindow

class ClosedContracts(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("დახურული ხელშეკრულებები")
        self.setWindowIcon(QIcon("Icons/closed_contracts.png"))
        self.resize(1400, 800)


        layout = QGridLayout()

        # --------------------------------------------Box1-----------------------------------------------------
        box1 = QGroupBox("ნავიგაცია")
        box1.setFixedWidth(275)
        box1.setStyleSheet("QGroupBox {font-style: italic; font-size: 10pt; }")
        box_layout = QGridLayout()


        # Export
        export = QToolButton()
        export.setText(" ექსპორტი ")
        export.setIcon(QIcon("Icons/excel_icon.png"))
        export.setIconSize(QSize(37, 40))
        export.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        export.setStyleSheet("font-size: 16px;")
        # export.clicked.connect(self.export_to_excel)

        # Checkbox
        # all_together = QCheckBox("ყველა ერთად")

        # BlankUnderTheCheckBox
        # blank_under_the_check_box = QLabel("")
        # blank_under_the_check_box.setStyleSheet("border: 1px solid gray; padding: 5px;")
        # blank_under_the_check_box.setFixedHeight(30)


        # # Box1 widgets
        # box_layout.addWidget(add_item, 0, 0)
        # box_layout.addWidget(edit_item, 0, 2)
        # box_layout.addWidget(choosing_the_field, 0, 2)
        box_layout.addWidget(export, 0, 3)
        # box_layout.addWidget(all_together, 1, 1)
        # box_layout.addWidget(blank_under_the_check_box, 1, 2)

        # Placing box1
        layout.addWidget(box1, 0, 0, 1, 1)
        box1.setLayout(box_layout)

        # --------------------------------------------Box2-----------------------------------------------------
        box2 = QGroupBox("ძებნა")
        box2.setStyleSheet("QGroupBox {font-style: italic; font-size: 10pt; }")
        box_layout2 = QGridLayout()

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("მოძებნე აქ")
        self.search_input.textChanged.connect(self.apply_text_filter)

        # Radio Button
        self.contract_radio = QRadioButton("ხელშეკრულების N")
        self.name_radio = QRadioButton("სახელი და გვარი")
        self.id_radio = QRadioButton("პირადი N")
        self.model_radio = QRadioButton("მოდელი")
        self.tel_radio = QRadioButton("ტელეფონის ნომერი")
        self.imei_radio = QRadioButton("IMEI")

        # Grouping buttons
        button_group = QButtonGroup()
        button_group.addButton(self.contract_radio)
        button_group.addButton(self.name_radio)
        button_group.addButton(self.id_radio)
        button_group.addButton(self.model_radio)
        button_group.addButton(self.tel_radio)
        button_group.addButton(self.imei_radio)

        # Making Widgets
        box_layout2.addWidget(self.contract_radio)
        box_layout2.addWidget(self.name_radio)
        box_layout2.addWidget(self.id_radio)
        box_layout2.addWidget(self.model_radio)
        box_layout2.addWidget(self.tel_radio)
        box_layout2.addWidget(self.imei_radio)
        box_layout2.addWidget(self.search_input, 2, 1)



        layout.addWidget(box2, 0, 1, 1, 1)
        box2.setLayout(box_layout2)

        # --------------------------------------------Box3-----------------------------------------------------
        box3 = QGroupBox("თარიღის მიხედვით გაფილტვრა")
        box3.setStyleSheet("QGroupBox {font-style: italic; font-size: 10pt; }")
        box_layout3 = QGridLayout()


        self.from_date = QDateEdit()
        self.from_date.setCalendarPopup(True)
        self.from_date.setDate(QDate.currentDate().addMonths(-1))
        self.to_date = QDateEdit()
        self.to_date.setCalendarPopup(True)
        self.to_date.setDate(QDate.currentDate())

        self.contract_date_radio = QRadioButton("გაფორმების თარიღით")
        self.contract_date_radio.setChecked(True)  # Default
        self.closing_date_radio = QRadioButton("დახურვის თარიღით")

        box_layout3.addWidget(self.contract_date_radio, 0, 0, 1, 2)
        box_layout3.addWidget(self.closing_date_radio, 1, 0, 1, 2)

        box_layout3.addWidget(QLabel("დან თარიღი:"), 2, 0)
        box_layout3.addWidget(self.from_date, 2, 1)
        box_layout3.addWidget(QLabel("მდე თარიღი:"), 3, 0)
        box_layout3.addWidget(self.to_date, 3, 1)

        filter_button = QPushButton("ძებნა")
        filter_button.clicked.connect(self.search_by_date)
        refresh_button = QPushButton("განახლება")
        refresh_button.clicked.connect(self.refresh_table)

        box_layout3.addWidget(refresh_button, 4, 0)
        box_layout3.addWidget(filter_button, 4, 1)

        layout.addWidget(box3, 0, 2, 1, 1)
        box3.setLayout(box_layout3)


        # --------------------------------------------Table-----------------------------------------------------
        self.db = QSqlDatabase.addDatabase("QSQLITE")
        self.db.setDatabaseName("Databases/closed_contracts.db")
        if not self.db.open():
            raise Exception("ბაზასთან კავშირი ვერ მოხერხდა")

        self.table = QTableView()
        self.model = QSqlTableModel(self, self.db)
        self.model.setTable("closed_contracts")
        self.model.select()
        self.table.setModel(self.model)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)  # Read-only table
        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setSelectionMode(QTableView.SingleSelection)

        layout.addWidget(self.table, 1, 0, 4, 5)


        # --------------------------------------------Layout-----------------------------------------------------
        self.setLayout(layout)


        # --------------------------------------------Functions-----------------------------------------------------

    def refresh_table(self):
        self.model.setFilter("")  # Clears filter
        self.model.select()  # This reloads the data from DB

    def search_by_date(self):
        from_date_str = self.from_date.date().toString("dd.MM.yyyy")
        to_date_str = self.to_date.date().toString("dd.MM.yyyy")

        if self.contract_date_radio.isChecked():
            date_column = "date"
        elif self.closing_date_radio.isChecked():
            date_column = "closing_date"
        else:
            QMessageBox.warning(self, "შეცდომა", "აირჩიეთ თარიღის ტიპი.")
            return

        # Assuming your table has a date column named 'date'
        filter_str = f"{date_column} >= '{from_date_str}' AND date <= '{to_date_str}'"
        self.model.setFilter(filter_str)
        self.model.select()

    def apply_text_filter(self, text):
        column = ""

        if self.contract_radio.isChecked():
            column = "id"
        elif self.name_radio.isChecked():
            column = "name_surname"
        elif self.id_radio.isChecked():
            column = "id_number"
        elif self.model_radio.isChecked():
            column = "model"
        elif self.tel_radio.isChecked():
            column = "tel_number"
        elif self.imei_radio.isChecked():
            column = "imei"

        if column:
            filter_str = f"{column} LIKE '%{text}%'"
            self.model.setFilter(filter_str)
        else:
            self.model.setFilter("")  # No filter if nothing selected

        self.model.select()



    def export_to_excel(self):

        row_count = self.model.rowCount()
        col_count = self.model.columnCount()

        headers = [self.model.headerData(col, Qt.Horizontal) for col in range(col_count)]
        data = [
            [self.model.data(self.model.index(row, col)) for col in range(col_count)]
            for row in range(row_count)
        ]

        df = pd.DataFrame(data, columns=headers)

        try:
            temp_path = os.path.join(tempfile.gettempdir(), "temp_export.xlsx")
            df.to_excel(temp_path, index=False)

            # Open Excel file
            os.startfile(temp_path)  # Safer and native on Windows
        except Exception as e:
            QMessageBox.critical(self, "შეცდომა", str(e))

