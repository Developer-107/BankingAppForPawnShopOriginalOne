import os
import sys
import urllib

import psycopg
from PyQt5.QtSql import QSqlDatabase
from dotenv import load_dotenv


load_dotenv()

office_mob_number = os.environ["OFFICE_MOB_NUMBER"]

def resource_path(relative_path):
    """
    Returns the absolute path relative to the .exe folder (or script folder),
    ignoring PyInstaller's temp folder. Use this if you want to keep all resources external.
    """
    base_path = os.path.dirname(sys.argv[0])
    return os.path.normpath(os.path.join(base_path, relative_path))

def get_conn():
    conn = psycopg.connect(os.environ["DATABASE_URL"])
    conn.autocommit = True
    return conn


def get_qt_db(unique_connection_name=""):
    url = urllib.parse.urlparse(os.environ["DATABASE_URL"])

    db = QSqlDatabase.addDatabase("QPSQL", unique_connection_name)
    db.setHostName(url.hostname)
    db.setPort(url.port or 5432)
    db.setDatabaseName(url.path[1:])  # remove leading '/'
    db.setUserName(url.username)
    db.setPassword(url.password)

    if not db.open():
        print(db.lastError().text())
        raise Exception(f"Cannot connect to DB: {db.lastError().text()}")

    return db


