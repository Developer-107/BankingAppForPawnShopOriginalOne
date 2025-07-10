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


class DetailWindow(QWidget):
    def __init__(self, contract_id, name_surname, item_name):
        super().__init__()
        self.setWindowTitle("დარიცხული და გადახდილი პროცენტები")
        self.setWindowIcon(QIcon("Icons/percent_payment_icon.png"))
        self.resize(1250, 540)
        self.contract_id = contract_id
        self.name_surname = name_surname
        self.item_name = item_name

        layout = QGridLayout()


        # --------------------------------------------ExportBox-----------------------------------------------------
        export_box = QGroupBox("ნავიგაცია")
        export_box.setFixedWidth(275)
        export_box.setStyleSheet("QGroupBox {font-style: italic; font-size: 10pt; }")
        box_layout = QGridLayout()

        # Export
        export_table = QToolButton()
        export_table.setText(" ექსპორტი ")
        export_table.setIcon(QIcon("Icons/excel_icon.png"))
        export_table.setIconSize(QSize(37, 40))
        export_table.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        export_table.setStyleSheet("font-size: 16px;")
        # export_table.clicked.connect(self.export_table_to_excel)

        box_layout.addWidget(export_table, 0, 3)
        # Placing box1
        layout.addWidget(export_box, 0, 0, 1, 1)
        export_box.setLayout(box_layout)

        # --------------------------------------------Box1-----------------------------------------------------

        box1 = QGroupBox("მონაცემები")
        box1.setStyleSheet("""
            QGroupBox {
                font-style: italic;
                font-size: 10pt;
                border: 1px solid gray;
                border-radius: 5px;
                margin-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 3px;
            }
            QLabel {
                font-size: 11pt;
                color: #2c3e50;
            }
        """)
        box_layout1 = QGridLayout()


        #
        contract_id_label = QLabel(f"ხელშეკრულების ნომერი N: {self.contract_id}")
        name_surname_label = QLabel(f"სახელი და გვარი: {self.name_surname}")
        item_name_label = QLabel(f"ნივთის დასახელება: {self.item_name}")

        box_layout1.addWidget(contract_id_label, 0, 0)
        box_layout1.addWidget(name_surname_label, 0, 1)
        box_layout1.addWidget(item_name_label, 0, 2)




        layout.addWidget(box1, 0, 1, 1, 3)
        box1.setLayout(box_layout1)


        # --------------------------------------------Table1-----------------------------------------------------
        name_table1 = QLabel("დარიცხული პროცენტები")
        name_table1.setStyleSheet("font-size: 16px; font-weight: bold;")

        layout.addWidget(name_table1, 1, 0)

        self.db1 = QSqlDatabase.addDatabase("QSQLITE", "adding_percent_amount")
        self.db1.setDatabaseName("Databases/adding_percent_amount.db")
        if not self.db1.open():
            raise Exception("ბაზასთან კავშირი ვერ მოხერხდა")

        self.table1 = QTableView()
        self.model1 = QSqlTableModel(self, self.db1)
        self.model1.setTable("adding_percent_amount")
        self.model1.select()
        self.table1.setModel(self.model1)
        self.table1.setEditTriggers(QAbstractItemView.NoEditTriggers)  # Read-only table
        self.table1.setSelectionBehavior(QTableView.SelectRows)
        self.table1.setSelectionMode(QTableView.SingleSelection)

        # Filter rows by contract_id
        self.model1.setFilter(f"contract_id = {contract_id}")
        self.model1.select()

        # Show only selected columns
        columns_to_show = ["contract_id", "date_of_percent_addition", "percent_amount"]
        for col in range(self.model1.columnCount()):
            field = self.model1.headerData(col, Qt.Horizontal)
            self.table1.setColumnHidden(col, field not in columns_to_show)

        self.table1.resizeColumnsToContents()

        layout.addWidget(self.table1, 2, 0, 1, 2)


        # --------------------------------------------Table2-----------------------------------------------------

        name_table2 = QLabel("გადახდილი პროცენტები")
        name_table2.setStyleSheet("font-size: 16px; font-weight: bold;")


        layout.addWidget(name_table2, 1, 2)

        self.db2 = QSqlDatabase.addDatabase("QSQLITE", "paid_percent_amount")
        self.db2.setDatabaseName("Databases/paid_percent_amount.db")
        if not self.db2.open():
            raise Exception("ბაზასთან კავშირი ვერ მოხერხდა")

        self.table2 = QTableView()
        self.model2 = QSqlTableModel(self, self.db2)
        self.model2.setTable("paid_percent_amount")
        self.model2.select()
        self.table2.setModel(self.model2)
        self.table2.setEditTriggers(QAbstractItemView.NoEditTriggers)  # Read-only table
        self.table2.setSelectionBehavior(QTableView.SelectRows)
        self.table2.setSelectionMode(QTableView.SingleSelection)

        # Filter rows by contract_id
        self.model2.setFilter(f"contract_id = {contract_id}")
        self.model2.select()

        # Show only selected columns
        columns_to_show = ["date_of_percent_addition", "paid_amount"]
        for col in range(self.model2.columnCount()):
            field = self.model2.headerData(col, Qt.Horizontal)
            self.table2.setColumnHidden(col, field not in columns_to_show)

        self.table2.resizeColumnsToContents()

        layout.addWidget(self.table2, 2, 2, 1, 2)

        # --------------------------------------------Layout-----------------------------------------------------
        self.setLayout(layout)


        # --------------------------------------------Functions-----------------------------------------------------

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
