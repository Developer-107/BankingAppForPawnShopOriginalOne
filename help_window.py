from PyQt5.QtCore import QSize, Qt
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QGroupBox


class HelpWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("დახმარება")
        self.setWindowIcon(QIcon("Icons/help.png"))
        self.resize(400, 100)

        label = QLabel("დახმარებისთვის მიმართეთ შესაბამის სამსახურს!")
        label.setAlignment(Qt.AlignCenter)  # Center inside layout

        layout = QVBoxLayout()
        layout.addWidget(label)

        self.setLayout(layout)

        self.setStyleSheet("""
                   QLabel {
                       font-style: Bold;
                       font-size: 9pt;
                   }
               """)
