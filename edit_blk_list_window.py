from PyQt5.QtCore import pyqtSignal

from utils import get_conn
from PyQt5.QtWidgets import QWidget, QGridLayout, QLabel, QLineEdit, QPushButton, QMessageBox
from PyQt5.QtGui import QIcon

from utils import resource_path


class EditBlkListWindow(QWidget):

    closing_signal = pyqtSignal()

    def __init__(self, record_id, parent=None):
        super().__init__(parent)
        self.record_id = record_id
        self.setWindowTitle("შავი სიის ჩანაწერის რედაქტირება")
        self.setWindowIcon(QIcon(resource_path("Icons/edit_icon.png")))
        self.resize(600, 300)
        self.build_ui()
        self.load_data()

    def build_ui(self):
        layout = QGridLayout()

        self.name_surname_box = QLineEdit()
        self.id_number_box = QLineEdit()
        self.tel_number_box = QLineEdit()
        self.imei_box = QLineEdit()

        layout.addWidget(QLabel("სახელი და გვარი:"), 0, 0)
        layout.addWidget(self.name_surname_box, 0, 1)
        layout.addWidget(QLabel("პირადი ნომერი:"), 1, 0)
        layout.addWidget(self.id_number_box, 1, 1)
        layout.addWidget(QLabel("ტელეფონის ნომერი:"), 2, 0)
        layout.addWidget(self.tel_number_box, 2, 1)
        layout.addWidget(QLabel("IMEI:"), 3, 0)
        layout.addWidget(self.imei_box, 3, 1)

        save_btn = QPushButton("შენახვა")
        save_btn.clicked.connect(self.save_changes)
        layout.addWidget(save_btn, 4, 1)

        self.setLayout(layout)

    def load_data(self):
        conn = get_conn()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM black_list WHERE id = %s", (self.record_id,))
        row = cursor.fetchone()
        conn.close()

        if row:
            self.name_surname_box.setText(row[1])
            self.id_number_box.setText(row[2])
            self.tel_number_box.setText(row[3])
            self.imei_box.setText(row[4])

    def save_changes(self):
        conn = get_conn()
        cursor = conn.cursor()
        try:
            cursor.execute("""
                UPDATE black_list SET 
                    name_surname = %s, 
                    id_number = %s, 
                    tel_number = %s, 
                    imei = %s
                WHERE id = %s
            """, (
                self.name_surname_box.text(),
                self.id_number_box.text(),
                self.tel_number_box.text(),
                self.imei_box.text(),
                self.record_id
            ))
            conn.commit()
            QMessageBox.information(self, "წარმატება", "ჩანაწერი წარმატებით განახლდა")
            self.close()
        except Exception as e:
            QMessageBox.critical(self, "შეცდომა", f"შეცდომა:\n{e}")
        finally:
            conn.close()



    def closeEvent(self, event):
        self.closing_signal.emit()
        event.accept()
