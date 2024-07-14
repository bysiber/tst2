import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path


LOG_FILE = Path.cwd() / "output" / "automation.log"

if not LOG_FILE.parent.exists():
    LOG_FILE.parent.mkdir()

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

console_handler = logging.StreamHandler()
file_handler = RotatingFileHandler(LOG_FILE, maxBytes=20971520, encoding="utf-8", backupCount=50)
console_handler.setLevel(logging.INFO)
file_handler.setLevel(logging.DEBUG)

console_log_format = "%(asctime)s [%(levelname)5s] %(lineno)3d: %(message)s"
file_log_format = "%(asctime)s [%(levelname)5s] %(filename)s:%(lineno)3d: %(message)s"
console_formatter = logging.Formatter(console_log_format, datefmt="%d-%m-%Y %H:%M:%S")
console_handler.setFormatter(console_formatter)
file_formatter = logging.Formatter(file_log_format, datefmt="%d-%m-%Y %H:%M:%S")
file_handler.setFormatter(file_formatter)

logger.addHandler(console_handler)
logger.addHandler(file_handler)

zeep_logger = logging.getLogger("zeep.transports")
zeep_logger.setLevel(logging.DEBUG)
zeep_logger.addHandler(console_handler)
zeep_logger.propagate = True
