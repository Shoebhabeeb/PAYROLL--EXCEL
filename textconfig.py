import configparser
import os
from collections import namedtuple


config = configparser.ConfigParser()
config_file = 'text_config.ini'
config.read(config_file)


def has_parent(path):

    par_dir = os.path.abspath(os.path.join(path, os.pardir))
    return par_dir, os.path.exists(par_dir)


def mkdir_if_parent_present(path):

    par_dir, present = has_parent(path)
    if present:
        if not os.path.exists(path):
            os.mkdir(path)
        return path
    else:
        raise FileNotFoundError(f'Please ensure that {par_dir} exists. Check {config_file}')


# text to excel
INPUT_TEXT_DIR = mkdir_if_parent_present(config['TEXT']['INPUT_TEXT_DIR'])
OUTPUT_EXCEL_DIR = mkdir_if_parent_present(config['TEXT']['OUTPUT_EXCEL_DIR'])
OUTPUT_ERROR_DIR = mkdir_if_parent_present(config['TEXT']['OUTPUT_ERROR_DIR'])
OUTPUT_ZIP_FILENAME = config['TEXT']['OUTPUT_ZIP_FILENAME']
OUTPUT_EXCEL_ZIP = mkdir_if_parent_present(config['TEXT']['OUTPUT_EXCEL_ZIP'])
OUTPUT_LOG_DIR = mkdir_if_parent_present(config['TEXT']['OUTPUT_LOG_DIR'])

text_tuple = namedtuple('TEXT', ['INPUT_TEXT_DIR',
                                 'OUTPUT_EXCEL_DIR',
                                 'OUTPUT_ERROR_DIR',
                                 'OUTPUT_ZIP_FILENAME',
                                 'OUTPUT_EXCEL_ZIP',
                                 'OUTPUT_LOG_DIR']
                        )
TEXT = text_tuple(INPUT_TEXT_DIR, OUTPUT_EXCEL_DIR, OUTPUT_ERROR_DIR,
                  OUTPUT_ZIP_FILENAME, OUTPUT_EXCEL_ZIP, OUTPUT_LOG_DIR)


__all__ = [TEXT]

if __name__ == "__main__":
    print(TEXT)

