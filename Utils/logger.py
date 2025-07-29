import inspect
import logging
import os
import sys
from logging.handlers import TimedRotatingFileHandler
from pathlib import Path

sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
curr_file_path = Path(__file__)
root_dir = curr_file_path.parent.parent.absolute()

class Logger:

    @staticmethod
    def logging_module(uniaue_id=None, flag=None, screenshot_path=None):
        # To get the name of testcaes file name at runtime
        logger = logging.getLogger(inspect.stack()[1][3])

        # create handler
        if not flag:
            handler = TimedRotatingFileHandler(filename = f"../Logs/{uniaue_id}.log", when = 'D',
                                               interval=1, backupCount=90,
                                               encoding='uft-8', delay=False)
        else:
            filename = os.path.join(screenshot_path, "../TestcaseLogs", f"{uniaue_id}.log")
            if not os.path.exists(filename):
                with open(filename, 'w'): #This will create an empty file if it doesn't exist
                    pass
            handler = TimedRotatingFileHandler(filename = filename, when = 'D',
                                               interval=1, backupCount=90,
                                               encoding='uft-8', delay=False)

        # handler = logging.FileHandler(filename=LOGGER_PATH)
        #create formatter and add to handler
        formatter = logging.Formatter("%(asctime)s : %(levelname)s : [%(filename)s: %(lineno)s- %(funcName)s()] : %(message)s", "%Y-%m-%d %H:%M:%S")
        handler.setFormatter(formatter)
        # add the handler to named logger
        logger.addHandler(handler)
        # set the logging level
        logger.setLevel(logging.INFO)
        return logger
