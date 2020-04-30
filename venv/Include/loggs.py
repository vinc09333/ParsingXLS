import logging
from logging.handlers import TimedRotatingFileHandler
import datetime
def log_info(msg):
    logfile = open(file='tmp/info.log', mode='a', encoding='utf-8')
    logfile.write(f'{datetime.datetime.today().strftime("[%Y-%m-%d|%H:%M]")} {msg}\n')
def log_error(msg):
    logfile = open(file='tmp/error.log', mode='a', encoding='utf-8')
    logfile.write(f'{datetime.datetime.today().strftime("[%Y-%m-%d|%H:%M]")} {msg}\n')
