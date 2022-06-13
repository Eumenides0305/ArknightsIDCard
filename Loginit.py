import logging
from logging import handlers


level_relations = {
    'debug': logging.DEBUG,
    'info': logging.INFO,
    'warning': logging.WARNING,
    'error': logging.ERROR,
    'crit': logging.CRITICAL
}


def _get_logger(filename, level='info'):
    log = logging.getLogger(filename)
    log.setLevel(level_relations.get(level))
    fmt = logging.Formatter('%(asctime)s %(thread)d %(filename)s[line:%(lineno)d] - %(levelname)s: %(message)s')
    file_handler = handlers.TimedRotatingFileHandler(filename=filename, when='D', backupCount=1, encoding='utf-8')
    file_handler.setFormatter(fmt)
    log.addHandler(file_handler)
    return log
