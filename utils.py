import os
import sys
import psycopg
from dotenv import load_dotenv


load_dotenv()

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