import os
from datetime import datetime


def is_file_write_locked(path):
    if not (os.path.exists(path)):
        return False

    try:
        f = open(path, 'ab')
        f.close()
    except IOError:
        return True


def today_as_string():
    return datetime.today().strftime("%Y.%m.%d")
