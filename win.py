import ctypes
import tkinter as tk
from tkinter import filedialog
import logging
from logging import handlers


def get_logger(filename, level='info'):
    level_relations = {
        'debug': logging.DEBUG,
        'info': logging.INFO,
        'warning': logging.WARNING,
        'error': logging.ERROR,
        'crit': logging.CRITICAL
    }
    log = logging.getLogger(filename)
    log.setLevel(level_relations.get(level))
    fmt = logging.Formatter('%(asctime)s %(thread)d %(filename)s[line:%(lineno)d] - %(levelname)s: %(message)s')
    file_handler = handlers.TimedRotatingFileHandler(filename=filename, when='D', backupCount=1, encoding='utf-8')
    file_handler.setFormatter(fmt)
    log.addHandler(file_handler)
    return log


def choose_excel(defaultOpenPath: str = './') -> str:
    root = tk.Tk()
    root.withdraw()
    f_path = ''
    tryTimes = 0
    while f_path == '' and tryTimes < 3:
        tryTimes = tryTimes + 1
        f_path = filedialog.askopenfilename(title='select excel file', initialdir=defaultOpenPath, filetypes=[('Excel', '.xls .xlsx')])
    return f_path


def info(message: str):
    ctypes.windll.user32.MessageBoxW(0, message, "ArknightsIDCard", 0 | 64)
