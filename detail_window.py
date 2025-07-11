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
        export_table.clicked.connect(self.export_both_tables_to_excel)

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

        record = self.model1.record()
        column_indices = {record.field(i).name(): i for i in range(record.count())}

        column_labels = {
            "contract_id": "ხელშეკრულების N",
            "date_of_percent_addition": "პროცენტის დამატების თარიღი",
            "percent_amount": "დარიცხული პროცენტი"
        }

        for name, label in column_labels.items():
            if name in column_indices:
                self.model1.setHeaderData(column_indices[name], Qt.Horizontal, label)

        # Make header text bold
        header = self.table1.horizontalHeader()
        font = header.font()
        font.setBold(True)
        header.setFont(font)
        header.setStyleSheet("""
               QHeaderView::section {
               padding: 4px 8px;
               }
                        """)

        self.table1.resizeColumnsToContents()

        layout.addWidget(self.table1, 2, 0, 1, 2)

        # --------------------------------------------SumQLabel--------------------------------------------------
        # Footer for table1
        self.left_footer_label = QLabel("სულ: 0.00 ₾")
        self.left_footer_label.setAlignment(Qt.AlignRight)
        self.left_footer_label.setStyleSheet("""
            background-color: yellow;
            font-weight: bold;
            padding: 6px;
            font-size: 14px;
        """)
        layout.addWidget(self.left_footer_label, 3, 0, 1, 2)  # Below table1

        # Footer for table2
        self.right_footer_label = QLabel("სულ: 0.00 ₾")
        self.right_footer_label.setAlignment(Qt.AlignRight)
        self.right_footer_label.setStyleSheet("""
            background-color: yellow;
            font-weight: bold;
            padding: 6px;
            font-size: 14px;
        """)
        layout.addWidget(self.right_footer_label, 3, 2, 1, 2)  # Below table2



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

        record = self.model2.record()
        column_indices = {record.field(i).name(): i for i in range(record.count())}

        column_labels = {
            "date_of_percent_addition": "პროცენტის გადახდის თარიღი",
            "paid_amount": "გადახილი პროცენტი"
        }

        for name, label in column_labels.items():
            if name in column_indices:
                self.model2.setHeaderData(column_indices[name], Qt.Horizontal, label)


        header = self.table2.horizontalHeader()
        font = header.font()
        font.setBold(True)
        header.setFont(font)
        header.setStyleSheet("""
                            QHeaderView::section {
                                padding: 4px 8px;
                            }
                        """)

        self.table2.resizeColumnsToContents()

        layout.addWidget(self.table2, 2, 2, 1, 2)

        # --------------------------------------------FunctionsForSum-----------------------------------------------------

        self.update_footer_sum(self.model1, self.left_footer_label, "percent_amount")
        self.update_footer_sum(self.model2, self.right_footer_label, "paid_amount")

        # --------------------------------------------Layout-----------------------------------------------------
        self.setLayout(layout)


        # --------------------------------------------Functions-----------------------------------------------------

    def export_both_tables_to_excel(self):
        try:
            # --- Table 1 ---
            columns_to_export1 = ["contract_id", "date_of_percent_addition", "percent_amount"]
            data1 = []

            for row in range(self.model1.rowCount()):
                row_data = []
                for col in range(self.model1.columnCount()):
                    header = self.model1.headerData(col, Qt.Horizontal)
                    if header in columns_to_export1:
                        value = self.model1.data(self.model1.index(row, col))
                        row_data.append(value)
                data1.append(row_data)

            df1 = pd.DataFrame(data1, columns=columns_to_export1)

            # --- Table 2 ---
            columns_to_export2 = ["date_of_percent_addition", "paid_amount"]
            data2 = []

            for row in range(self.model2.rowCount()):
                row_data = []
                for col in range(self.model2.columnCount()):
                    header = self.model2.headerData(col, Qt.Horizontal)
                    if header in columns_to_export2:
                        value = self.model2.data(self.model2.index(row, col))
                        row_data.append(value)
                data2.append(row_data)

            df2 = pd.DataFrame(data2, columns=columns_to_export2)

            # --- Write both tables side by side in Excel ---
            temp_path = os.path.join(tempfile.gettempdir(), "percent_export.xlsx")

            with pd.ExcelWriter(temp_path, engine='openpyxl') as writer:
                df1.to_excel(writer, sheet_name='Details', index=False, startrow=0, startcol=0)
                df2.to_excel(writer, sheet_name='Details', index=False, startrow=0, startcol=len(df1.columns))

            os.startfile(temp_path)

        except Exception as e:
            QMessageBox.critical(self, "შეცდომა", str(e))

    def update_footer_sum(self, model, footer_label, column_name):
        total = 0.0
        record = model.record()
        col_index = next((i for i in range(record.count()) if record.field(i).name() == column_name), -1)

        if col_index != -1:
            for row in range(model.rowCount()):
                index = model.index(row, col_index)
                value = model.data(index)
                try:
                    total += float(value)
                except (TypeError, ValueError):
                    pass

        footer_label.setText(f"სულ: {total:.2f} ₾")