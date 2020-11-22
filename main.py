import logging
import time
import os
import traceback

import win32com
import win32com.client

import Setup as S

def main():
    logger = logging.getLogger('Моя программа')
    logger.setLevel(logging.INFO)
    if not os.path.isdir(S.LOG_PATH):
        os.makedirs(S.LOG_PATH)
    f_handler = logging.FileHandler(S.LOG)
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    f_handler.setFormatter(formatter)
    logger.addHandler(f_handler)

    logger.info('Старт программы')

    while True:
        try:
            time.sleep(5)
        except:
            logger.error('Critical error', exc_info = True)
    
if __name__ == '__main__':
    main()