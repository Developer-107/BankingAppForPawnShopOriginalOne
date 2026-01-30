from PyQt5.QtCore import QDateTime, QTimer
from PyQt5.QtWidgets import QMessageBox, QLabel

from load_data_worker import LoadDataWorker

# --------------------------
# Daily Timer Setup
# --------------------------
def schedule_daily_load(window):
    """Schedules LoadDataWorker to run at the next midnight (00:00)"""
    now = QDateTime.currentDateTime()
    # Next midnight
    midnight = QDateTime(now.date().addDays(1))
    midnight.setTime(QDateTime().time().fromString("00:00:00", "hh:mm:ss"))

    msecs_until_midnight = now.msecsTo(midnight)

    window.timer = QTimer(window)
    window.timer.setSingleShot(True)
    # Use lambda to pass window to function
    window.timer.timeout.connect(lambda: run_and_reschedule(window))
    window.timer.start(msecs_until_midnight)


def run_and_reschedule(window):
    """Runs the loader and schedules it for the next day"""
    start_load_data(window)
    schedule_daily_load(window)  # schedule for next day


# --------------------------
# Worker Setup
# --------------------------
def start_load_data(window):
    """Starts the LoadDataWorker if not already running"""
    if hasattr(window, 'worker') and window.worker.isRunning():
        return

    window.worker = LoadDataWorker()
    # Connect signals with lambdas to pass the window
    window.worker.error.connect(lambda msg: on_load_error(window, msg))
    window.worker.row_error.connect(lambda row, msg: on_row_error(window, row, msg))
    window.worker.finished.connect(lambda: on_load_success(window))
    window.worker.start()


# --------------------------
# Error Handlers
# --------------------------
def on_load_success(window):
    """Shows a temporary, non-blocking success message on the window."""
    toast = QLabel("ბაზა წარმატებით განახლდა!", window)
    toast.setStyleSheet("""
            background-color: #28a745;  /* green */
            color: white;
            padding: 8px;
            border-radius: 5px;
            font-weight: bold;
        """)
    toast.adjustSize()

    # Position it at top-right of the window
    toast.move(window.width() - toast.width() - 20, 20)

    toast.show()

    # Hide automatically after duration (ms)
    QTimer.singleShot(3000, toast.hide)


def on_load_error(window, msg):
    """Show a critical error message"""
    QMessageBox.critical(window, "Error", msg)


def on_row_error(window, row_number, message):
    """Show a warning for a specific row error"""
    QMessageBox.warning(
        window,
        "Row Error",
        f"Row {row_number} caused an error:\n{message}"
    )
