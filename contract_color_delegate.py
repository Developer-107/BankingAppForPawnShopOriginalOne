from PyQt5.QtWidgets import QStyledItemDelegate
from PyQt5.QtGui import QBrush, QColor
from PyQt5.QtCore import Qt, QDate

class ContractColorDelegate(QStyledItemDelegate):
    def paint(self, painter, option, index):
        model = index.model()
        row = index.row()

        try:
            rec = model.record(row)
            date_str = rec.value("date")
            day_quantity_raw = rec.value("day_quantity")
            percent_should_be_paid = float(rec.value("percent_should_be_paid") or 0)
            paid_percents = float(rec.value("paid_percents") or 0)
            principal_should_be_paid = float(rec.value("principal_should_be_paid") or 0)

            # ✅ Make green if nothing to pay
            if principal_should_be_paid == 0 and percent_should_be_paid == 0:
                bg_color = QColor("#2ff32f")
            else:
                contract_date = QDate.fromString(str(date_str).split(" ")[0], "yyyy-MM-dd")
                today = QDate.currentDate()

                if contract_date.isValid() and day_quantity_raw:
                    day_quantity = int(day_quantity_raw)
                    if day_quantity <= 0:
                        bg_color = QColor("white")
                    else:
                        days_diff = contract_date.daysTo(today)
                        is_payment_day = (days_diff + 1 >= day_quantity and (days_diff + 1) % day_quantity == 0)

                        if percent_should_be_paid > 0:
                            if is_payment_day:
                                bg_color = QColor("#f3f32f")  # due today & unpaid
                            elif days_diff > day_quantity:
                                bg_color = QColor("#f32f2f")     # overdue
                            else:
                                bg_color = QColor("white")
                        else:
                            bg_color = QColor("white")
                else:
                    bg_color = QColor("white")

        except Exception as e:
            print(f"Delegate error: {e}")
            bg_color = QColor("white")

        # Paint background
        painter.save()
        painter.fillRect(option.rect, bg_color)
        super().paint(painter, option, index)
        painter.restore()
