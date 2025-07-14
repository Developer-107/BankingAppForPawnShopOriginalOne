import os
import sys

def resource_path(relative_path):
    """
    Returns the absolute path relative to the .exe folder (or script folder),
    ignoring PyInstaller's temp folder. Use this if you want to keep all resources external.
    """
    base_path = os.path.dirname(sys.argv[0])
    return os.path.normpath(os.path.join(base_path, relative_path))