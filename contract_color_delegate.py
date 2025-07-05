from PyQt5.QtWidgets import QStyledItemDelegate
from PyQt5.QtGui import QBrush, QColor
from PyQt5.QtCore import Qt, QDate

class ContractColorDelegate(QStyledItemDelegate):
    def paint(self, painter, option, index):
        model = index.model()
        row = index.row()

        # Extract needed data from model:
        # Assuming these columns exist in your model:
        date_str = model.data(model.index(row, model.fieldIndex("date")))
        day_quantity = model.data(model.index(row, model.fieldIndex("day_quantity")))
        percent_should_be_paid = model.data(model.index(row, model.fieldIndex("percent_should_be_payed")))

        # Parse date string to QDate
        contract_date = QDate.fromString(date_str.split(" ")[0], "dd.MM.yyyy") if date_str else QDate()

        today = QDate.currentDate()
        start_date = min(contract_date, contract_date.addDays(int(day_quantity) - 1)) if contract_date.isValid() else QDate()

        # Default color
        bg_color = QColor("white")

        if contract_date.isValid() and day_quantity and percent_should_be_paid:
            try:
                day_quantity = int(day_quantity)
                percent_should_be_paid = float(percent_should_be_paid)
                days_diff = start_date.daysTo(today)

                if days_diff % day_quantity == 0 and percent_should_be_paid > 0:
                    bg_color = QColor("yellow")
                elif percent_should_be_paid > 0:
                    bg_color = QColor("red")
            except Exception:
                pass  # fallback to white if anything fails

        option.backgroundBrush = QBrush(bg_color)

        super().paint(painter, option, index)
