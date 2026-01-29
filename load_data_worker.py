from PyQt5.QtCore import QThread, pyqtSignal, QDateTime
from datetime import date

from utils import get_conn

class LoadDataWorker(QThread):
    finished = pyqtSignal()
    error = pyqtSignal(str)
    row_error = pyqtSignal(int, str)  # emits row index and error message

    def run(self):
        try:
            conn = get_conn()
            cursor = conn.cursor()

            cursor.execute("SELECT * FROM active_contracts_view")
            rows = cursor.fetchall()

            today = date.today()

            for row_index, row in enumerate(rows):
                try:

                    contract_id = row[0]
                    full_date_str = row[1]

                    contract_datetime = QDateTime.fromString(full_date_str, "yyyy-MM-dd HH:mm:ss")
                    if not contract_datetime.isValid():
                        raise ValueError(f"Invalid datetime format: {full_date_str}")
                    contract_date = contract_datetime.date()

                    name_surname = str(row[3])
                    id_number = str(row[4])
                    tel_number = str(row[5])
                    item_name = str(row[6])
                    model = str(row[7])
                    imei_sn = str(row[8])

                    # These could crash if values are None or strings
                    day_quantity = int(row[14])
                    percent = int(row[13])
                    principal_should_be_paid = float(row[17])
                    added_percents = float(row[18])
                    paid_percents = float(row[19])
                    percent_should_be_paid = float(row[20])

                    days_after = contract_date.daysTo(today)
                    new_added_percents = added_percents
                    first_due_day = contract_date.addDays(day_quantity - 1)
                    start_date = min(contract_date, first_due_day)

                    if days_after >= day_quantity > 0 and principal_should_be_paid > 0 and percent > 0:
                        days_diff = start_date.daysTo(today)

                        periods_passed = days_after // day_quantity
                        total_expected_adds = periods_passed + 1

                        already_added_times = self.get_already_added_times(cursor, contract_id)
                        additions_needed = total_expected_adds - already_added_times

                        if additions_needed > 0:
                            one_period_amount = (principal_should_be_paid * percent) / 100
                            new_added_percents = added_percents
                            status_for_added_percent = "დარიცხული პროცენტი"

                            for i in range(additions_needed):
                                new_added_percents += one_period_amount
                                cursor.execute("""
                                           INSERT INTO adding_percent_amount (
                                               contract_id, date_of_C_O, name_surname, id_number,
                                               tel_number, item_name, model, IMEI,
                                               date_of_percent_addition, percent_amount, status
                                           ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                                       """, (
                                    contract_id, full_date_str, name_surname, id_number,
                                    tel_number, item_name, model, imei_sn,
                                    QDateTime.currentDateTime().toString("yyyy-MM-dd HH:mm:ss"), one_period_amount,
                                    status_for_added_percent
                                ))

                    # Update DB
                    self.update_days_and_percents(cursor, contract_id, days_after, new_added_percents)


                except Exception as e:
                    self.row_error.emit(row_index + 1, str(e))

            conn.commit()
            conn.close()

            self.finished.emit()

        except Exception as e:
            self.error.emit(str(e))


    def update_days_and_percents(self, cursor, contract_id, days_after, new_added_percents):

        cursor.execute("""
            UPDATE active_contracts
            SET days_after_C_O = %s, added_percents = %s
            WHERE id = %s
        """, (days_after, new_added_percents, contract_id))

    @staticmethod
    def get_already_added_times(cursor, contract_id):
        cursor.execute("SELECT COUNT(*) FROM adding_percent_amount WHERE contract_id = %s",
                       (contract_id,))
        return cursor.fetchone()[0]
