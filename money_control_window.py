import os
import sqlite3
import tempfile

import pandas as pd
from PyQt5.QtCore import QDate, QSize, Qt
from PyQt5.QtGui import QIcon
from PyQt5.QtSql import QSqlDatabase, QSqlTableModel
from PyQt5.QtWidgets import QWidget, QGroupBox, QGridLayout, QDateEdit, QLabel, QPushButton, QTableView, \
    QAbstractItemView, QToolButton, QMessageBox
from openpyxl.chart import layout


class MoneyControl(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("თანხების კონეტროლი")
        self.setWindowIcon(QIcon("Icons/money_control.png"))
        self.resize(1400, 701)


        layout = QGridLayout()

        # --------------------------------------------Box1-----------------------------------------------------

        box1 = QGroupBox("თარიღის მიხედვით გაფილტვრა")
        box1.setStyleSheet("QGroupBox {font-style: italic; font-size: 10pt; }")
        box_layout1 = QGridLayout()

        self.from_date = QDateEdit()
        self.from_date.setCalendarPopup(True)
        self.from_date.setDate(QDate.currentDate().addMonths(-1))
        self.to_date = QDateEdit()
        self.to_date.setCalendarPopup(True)
        self.to_date.setDate(QDate.currentDate())
        #
        box_layout1.addWidget(QLabel("*დან თარიღი:"))
        box_layout1.addWidget(self.from_date)
        box_layout1.addWidget(QLabel("*მდე თარიღი:"))
        box_layout1.addWidget(self.to_date)

        filter_button = QPushButton("ძებნა")
        filter_button.clicked.connect(self.search_by_date)
        refresh_button = QPushButton("განახლება")
        refresh_button.clicked.connect(self.refresh_table)

        box_layout1.addWidget(refresh_button)
        box_layout1.addWidget(filter_button)

        layout.addWidget(box1, 0, 0, 1, 4)
        box1.setLayout(box_layout1)


        # --------------------------------------------Table1-----------------------------------------------------
        # Export
        export_table1 = QToolButton()
        export_table1.setText(" ექსპორტი ")
        export_table1.setIcon(QIcon("Icons/excel_icon.png"))
        export_table1.setIconSize(QSize(37, 40))
        export_table1.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        export_table1.setStyleSheet("font-size: 16px;")
        export_table1.clicked.connect(self.export_table1_to_excel)

        name_table1 = QLabel("შემოტანილი ძირი თანხა და პროცენტი")
        name_table1.setStyleSheet("font-size: 16px; font-weight: bold;")

        layout.addWidget(export_table1, 1, 0)
        layout.addWidget(name_table1, 1, 1)

        self.db1 = QSqlDatabase.addDatabase("QSQLITE", "paid_principle_and_paid_percentage_database")
        self.db1.setDatabaseName("Databases/paid_principle_and_paid_percentage_database.db")
        if not self.db1.open():
            raise Exception("ბაზასთან კავშირი ვერ მოხერხდა")

        self.table1 = QTableView()
        self.model1 = QSqlTableModel(self, self.db1)
        self.model1.setTable("paid_principle_and_paid_percentage_database")
        self.model1.select()
        self.table1.setModel(self.model1)
        self.table1.setEditTriggers(QAbstractItemView.NoEditTriggers)  # Read-only table
        self.table1.setSelectionBehavior(QTableView.SelectRows)
        self.table1.setSelectionMode(QTableView.SingleSelection)

        self.table1.resizeColumnsToContents()

        layout.addWidget(self.table1, 2, 0, 1, 2)


        # --------------------------------------------Table2-----------------------------------------------------
        # Export
        export_table2 = QToolButton()
        export_table2.setText(" ექსპორტი ")
        export_table2.setIcon(QIcon("Icons/excel_icon.png"))
        export_table2.setIconSize(QSize(37, 40))
        export_table2.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        export_table2.setStyleSheet("font-size: 16px;")
        export_table2.clicked.connect(self.export_table2_to_excel)

        name_table2 = QLabel("გაცემული და დამატებული თანხები")
        name_table2.setStyleSheet("font-size: 16px; font-weight: bold;")

        layout.addWidget(export_table2, 1, 2)
        layout.addWidget(name_table2, 1, 3)

        self.db2 = QSqlDatabase.addDatabase("QSQLITE", "given_and_additional_database")
        self.db2.setDatabaseName("Databases/given_and_additional_database.db")
        if not self.db2.open():
            raise Exception("ბაზასთან კავშირი ვერ მოხერხდა")

        self.table2 = QTableView()
        self.model2 = QSqlTableModel(self, self.db2)
        self.model2.setTable("given_and_additional_database")
        self.model2.select()
        self.table2.setModel(self.model2)
        self.table2.setEditTriggers(QAbstractItemView.NoEditTriggers)  # Read-only table
        self.table2.setSelectionBehavior(QTableView.SelectRows)
        self.table2.setSelectionMode(QTableView.SingleSelection)

        self.table2.resizeColumnsToContents()
        layout.addWidget(self.table2, 2, 2, 1, 2)

        # --------------------------------------------Layout-----------------------------------------------------
        self.setLayout(layout)


        # --------------------------------------------Functions-----------------------------------------------------
    def refresh_table(self):
        self.model1.setFilter("")  # Clears filter
        self.model1.select()  # This reloads the data from DB

        self.model2.setFilter("")
        self.model2.select()

    def search_by_date(self):
        from_date_str = self.from_date.date().toString("yyyy-MM-dd HH:mm:ss")
        to_date_str = self.to_date.date().toString("yyyy-MM-dd HH:mm:ss")

        date_column = "date_of_inflow"
        date_column1 = "date_of_outflow"

        # Assuming your table has a date column named 'date'
        filter_str = f"{date_column} >= '{from_date_str}' AND {date_column} <= '{to_date_str}'"
        self.model1.setFilter(filter_str)
        self.model1.select()

        filter_str1 = f"{date_column1} >= '{from_date_str}' AND {date_column1} <= '{to_date_str}'"
        self.model2.setFilter(filter_str1)
        self.model2.select()

    def export_table1_to_excel(self):

        row_count = self.model1.rowCount()
        col_count = self.model1.columnCount()

        headers = [self.model1.headerData(col, Qt.Horizontal) for col in range(col_count)]
        data = [
            [self.model1.data(self.model1.index(row, col)) for col in range(col_count)]
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


    def export_table2_to_excel(self):

        row_count = self.model2.rowCount()
        col_count = self.model2.columnCount()

        headers = [self.model2.headerData(col, Qt.Horizontal) for col in range(col_count)]
        data = [
            [self.model2.data(self.model2.index(row, col)) for col in range(col_count)]
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