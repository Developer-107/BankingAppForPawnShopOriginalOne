from utils import get_conn

from PyQt5.QtCore import QDate, QSize, Qt
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QWidget, QGridLayout, QLabel, QLineEdit, QToolButton, QHBoxLayout, QMessageBox

from utils import resource_path


class AddBlkListWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("დამატება")
        self.setWindowIcon(QIcon(resource_path("Icons/blacklist.png")))
        self.resize(400, 100)

        layout = QGridLayout()

        name_surname = QLabel("სახელი და გვარი:")
        id_number = QLabel("პირადი ნომერი:")
        tel_number = QLabel("საკონტაქტო ტელეფონი:")
        imei_sn = QLabel("IMEI:")

        self.name_surname_box = QLineEdit()
        self.name_surname_box.setPlaceholderText("სახელი და გვარი")

        self.id_number_box = QLineEdit()
        self.id_number_box.setPlaceholderText("პირადი ნომერი")

        self.tel_number_box = QLineEdit()
        self.tel_number_box.setPlaceholderText("საკონტაქტო ტელეფონი")

        self.imei_sn_box = QLineEdit()
        self.imei_sn_box.setPlaceholderText("IMEI")


        layout.addWidget(name_surname, 0, 0)
        layout.addWidget(id_number, 1, 0)
        layout.addWidget(tel_number, 2, 0)
        layout.addWidget(imei_sn, 3, 0)


        layout.addWidget(self.name_surname_box, 0, 1)
        layout.addWidget(self.id_number_box, 1, 1)
        layout.addWidget(self.tel_number_box, 2, 1)
        layout.addWidget(self.imei_sn_box, 3, 1)

        save_button = QToolButton()
        save_button.setText(" შენახვა ")
        save_button.setIcon(QIcon(resource_path("Icons/save_icon.png")))
        save_button.setIconSize(QSize(35, 35))
        save_button.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        save_button.setStyleSheet("font-size: 16px;")
        save_button.clicked.connect(self.save_to_blk_list_sql)

        close_window = QToolButton()
        close_window.setText(" დახურვა ")
        close_window.setIcon(QIcon(resource_path("Icons/cancel_icon.png")))
        close_window.setIconSize(QSize(35, 35))
        close_window.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        close_window.setStyleSheet("font-size: 16px;")
        close_window.clicked.connect(self.close)

        # Create horizontal layout for both buttons
        button_container = QWidget()
        button_layout = QHBoxLayout()
        button_layout.setContentsMargins(0, 0, 0, 0)  # remove padding
        button_layout.setSpacing(10)  # space between buttons

        button_layout.addWidget(save_button)
        button_layout.addWidget(close_window)
        button_container.setLayout(button_layout)

        # Add to main layout at row 5, column 3
        layout.addWidget(button_container, 4, 1)

        # --------------------------------------------Layout-----------------------------------------------------

        self.setLayout(layout)


    def save_to_blk_list_sql(self):
        try:
            conn = get_conn()
            cursor = conn.cursor()

            cursor.execute("""
                    INSERT INTO black_list (
                        name_surname, id_number, tel_number, imei
                    ) VALUES (%s, %s, %s, %s)
                """, (
                self.name_surname_box.text(),
                self.id_number_box.text(),
                self.tel_number_box.text(),
                self.imei_sn_box.text()
                ))

            conn.commit()
            QMessageBox.information(self, "წარმატება", "მონაცემები შენახულია")
            self.close()
        except Exception as e:
            QMessageBox.critical(self, "შეცდომა", f"ვერ შევინახე მონაცემები:\n{e}")
        finally:
            conn.close()



