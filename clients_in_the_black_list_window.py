import os
import tempfile

import pandas as pd
from PyQt5.QtCore import QSize, Qt
from PyQt5.QtGui import QIcon
from PyQt5.QtSql import QSqlDatabase, QSqlTableModel
from PyQt5.QtWidgets import QWidget, QTableView, QGridLayout, QToolButton, QGroupBox, QLineEdit, QRadioButton, \
    QButtonGroup, QAbstractItemView, QMessageBox

from edit_blk_list_window import EditBlkListWindow
from add_blk_list_window import AddBlkListWindow
from utils import resource_path, get_qt_db


class ClientsInTheBlackList(QWidget):
    def __init__(self, role):
        super().__init__()
        self.setWindowTitle("შავ სიაში მყოფი კლიენტები")
        self.setWindowIcon(QIcon(resource_path("Icons/blacklist.png")))
        self.resize(900, 500)
        self.role = role

        layout = QGridLayout()



        # --------------------------------------------Table-----------------------------------------------------

        self.db = get_qt_db()

        self.table = QTableView()
        self.model = QSqlTableModel(self, self.db)
        self.model.setTable("black_list")
        self.model.select()
        # Block to rename columns
        record = self.model.record()
        column_indices = {record.field(i).name(): i for i in range(record.count())}

        column_labels = {
            "id": "უნიკალური ნომერი N",
            "name_surname": "სახელი და გვარი",
            "id_number": "პირადობის ნომერი",
            "tel_number": "ტელეფონის ნომერი",
            "imei": "IMEI"
        }

        for name, label in column_labels.items():
            if name in column_indices:
                self.model.setHeaderData(column_indices[name], Qt.Horizontal, label)


        self.table.setModel(self.model)
        # Make header text bold
        header = self.table.horizontalHeader()
        font = header.font()
        font.setBold(True)
        header.setFont(font)
        header.setStyleSheet("""
                            QHeaderView::section {
                                padding: 4px 8px;
                            }
                        """)
        # Continue as usual
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)  # Read-only table
        self.table.setSelectionBehavior(QTableView.SelectRows)


        self.table.resizeColumnsToContents()
        layout.addWidget(self.table, 1, 1, 4, 5)

        # --------------------------------------------Box1-----------------------------------------------------

        box1 = QGroupBox("ნავიგაცია")
        box1.setStyleSheet("QGroupBox {font-style: italic; font-size: 10pt; }")
        box_layout = QGridLayout()

        # Box 1 buttons
        # Add
        add_to_blk_list = QToolButton()
        add_to_blk_list.setText(" დამატება ")
        # add_to_blk_list.setIcon(QIcon("Icons/.png"))
        add_to_blk_list.setFixedWidth(145)
        add_to_blk_list.setFixedHeight(45)
        add_to_blk_list.setIconSize(QSize(37, 40))
        add_to_blk_list.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        add_to_blk_list.setStyleSheet("font-size: 16px;")
        add_to_blk_list.clicked.connect(self.open_add_blk_list_window)

        # Edit
        edit_blk_list = QToolButton()
        edit_blk_list.setText(" რედაქტირება ")
        # edit_blk_list.setIcon(QIcon("Icons/.png"))
        edit_blk_list.setIconSize(QSize(37, 40))
        edit_blk_list.setFixedHeight(45)
        edit_blk_list.setFixedWidth(145)
        edit_blk_list.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        edit_blk_list.setStyleSheet("font-size: 16px;")
        edit_blk_list.clicked.connect(self.open_edit_blk_list_window)
        if self.role != "admin":
            edit_blk_list.setEnabled(False)

        # Remove
        delete_from_blk_list = QToolButton()
        delete_from_blk_list.setText(" ამოშლა ")
        # delete_from_blk_list.setIcon(QIcon("Icons/.png"))
        delete_from_blk_list.setIconSize(QSize(37, 40))
        delete_from_blk_list.setFixedHeight(45)
        delete_from_blk_list.setFixedWidth(145)
        delete_from_blk_list.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        delete_from_blk_list.setStyleSheet("font-size: 16px;")
        delete_from_blk_list.clicked.connect(self.delete_selected_row)
        if self.role != "admin":
            delete_from_blk_list.setEnabled(False)

        # Export
        export_blk_list = QToolButton()
        export_blk_list.setText(" ექსპორტი ")
        export_blk_list.setIcon(QIcon("Icons/excel_icon.png"))
        export_blk_list.setIconSize(QSize(37, 40))
        export_blk_list.setFixedHeight(45)
        export_blk_list.setFixedWidth(145)
        export_blk_list.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        export_blk_list.setStyleSheet("font-size: 16px;")
        export_blk_list.clicked.connect(self.export_blk_list_to_excel)

        # Box1 widgets
        box_layout.addWidget(add_to_blk_list, 0, 0)
        box_layout.addWidget(edit_blk_list, 1, 0)
        box_layout.addWidget(delete_from_blk_list, 2, 0)
        box_layout.addWidget(export_blk_list, 3, 0)

        layout.addWidget(box1, 1, 0)
        box1.setLayout(box_layout)


        # --------------------------------------------Box2-----------------------------------------------------
        box2 = QGroupBox("ძებნა")
        box2.setStyleSheet("QGroupBox {font-style: italic; font-size: 10pt; }")
        box_layout2 = QGridLayout()

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("მოძებნე აქ")
        self.search_input.textChanged.connect(self.apply_text_filter)

        self.name_radio = QRadioButton("სახელი და გვარი")
        self.id_radio = QRadioButton("პირადი N")
        self.tel_radio = QRadioButton("ტელეფონის ნომერი")
        self.imei_radio = QRadioButton("IMEI")

        # Grouping buttons
        button_group = QButtonGroup()
        button_group.addButton(self.name_radio)
        button_group.addButton(self.id_radio)
        button_group.addButton(self.tel_radio)
        button_group.addButton(self.imei_radio)

        # Making Widgets
        box_layout2.addWidget(self.name_radio)
        box_layout2.addWidget(self.id_radio)
        box_layout2.addWidget(self.tel_radio)
        box_layout2.addWidget(self.imei_radio)
        box_layout2.addWidget(self.search_input, 2, 1)

        layout.addWidget(box2, 0, 1, 1, 1)
        box2.setLayout(box_layout2)



        # --------------------------------------------Layout-----------------------------------------------------


        self.setLayout(layout)

    def open_add_blk_list_window(self):
        self.open_add_blk_list_window = AddBlkListWindow()
        self.open_add_blk_list_window.show()


    def apply_text_filter(self, text):
        column = ""

        if self.name_radio.isChecked():
            column = "name_surname"
        elif self.id_radio.isChecked():
            column = "id_number"
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

    def delete_selected_row(self):
        selected = self.table.selectionModel().selectedRows()
        if not selected:
            QMessageBox.warning(self, "შეცდომა", "გთხოვთ აირჩიოთ ერთი ჩანაწერი ამოსაშლელად")
            return

        row_index = selected[0].row()
        record_id = self.model.data(self.model.index(row_index, self.model.fieldIndex("id")))

        reply = QMessageBox.question(
            self,
            "დადასტურება",
            f"ნამდვილად გსურთ ჩანაწერის ამოშლა? ID: {record_id}",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self.model.removeRow(row_index)
            if self.model.submitAll():
                QMessageBox.information(self, "წარმატება", "ჩანაწერი ამოშლილია")
            else:
                QMessageBox.critical(self, "შეცდომა", "ჩანაწერის ამოშლა ვერ მოხერხდა")

    def open_edit_blk_list_window(self):
        selected = self.table.selectionModel().selectedRows()
        if not selected:
            QMessageBox.warning(self, "შეცდომა", "გთხოვთ აირჩიოთ ერთი ჩანაწერი რედაქტირებისთვის")
            return

        row = selected[0].row()
        record_id = self.model.data(self.model.index(row, self.model.fieldIndex("id")))
        self.edit_window = EditBlkListWindow(record_id)
        self.edit_window.show()
        self.edit_window.destroyed.connect(lambda: self.model.select())  # Refresh after edit

    def export_blk_list_to_excel(self):

        row_count = self.model.rowCount()
        col_count = self.model.columnCount()

        headers = [self.model.headerData(col, Qt.Horizontal) for col in range(col_count)]
        data = [
            [self.model.data(self.model.index(row, col)) for col in range(col_count)]
            for row in range(row_count)
        ]

        df = pd.DataFrame(data, columns=headers)

        try:
            temp_path = os.path.join(tempfile.gettempdir(), "temp_export_blk_list_to_excel_file.xlsx")
            df.to_excel(temp_path, index=False)

            # Open Excel file
            os.startfile(temp_path)  # Safer and native on Windows
        except Exception as e:
            QMessageBox.critical(self, "შეცდომა", str(e))
