import logging
import os
from datetime import datetime

def configure_logger(subfolder_name):
    current_directory = os.getcwd()
    logs_folder = os.path.join(current_directory, 'logs')

    if not os.path.exists(logs_folder):
        os.makedirs(logs_folder)

    subfolder_path = os.path.join(logs_folder, subfolder_name)

    if not os.path.exists(subfolder_path):
        os.makedirs(subfolder_path)

    current_date = datetime.now().strftime('%Y-%m-%d')
    log_file = os.path.join(subfolder_path, f"{current_date}.log")
    
    logger = logging.getLogger(subfolder_name)
    logger.setLevel(logging.INFO)
    
    file_handler = logging.FileHandler(log_file, mode='a')
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    
    return logger