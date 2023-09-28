import logging
import time
from logging.handlers import RotatingFileHandler

logger = logging.getLogger("AppLog")


def init_logging(level):
    """
    Creates a rotating log
    """
    logger.setLevel(level)

    handler = RotatingFileHandler("logs/automate.log", maxBytes=1024*1024, backupCount=5)
    formatter = logging.Formatter(
        '%(asctime)s [%(process)d]: %(message)s',
        '%b %d %H:%M:%S')
    formatter.converter = time.gmtime
    handler.setFormatter(formatter)
    logger.addHandler(handler)


def debug(message):
    logger.debug(message)


def error(message):
    logger.error(str(message))