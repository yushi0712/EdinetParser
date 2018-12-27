# xbrl から読み込む
import os
import sys

sys.path.append(os.path.join(os.path.dirname(__file__), 'XbrlReader'))


XBRL_ROOT_PATH = "C:/local/users/nishimura/Documents/!Theme/財務分析"
ORIGINAL_XBRL_DIR_PATH = "D:\Archive\財務データ\XBRL"
RENAMED_XBRL_DIR_NAME = "RenamedXBRL"
EDINET_INFO_FILE_NAME = "EDINETコード.xlsx"
XBRL_CONTENTS_FILE_NAME = "XBRL_Contents.xlsx"
ASR_SUMMARY_FILE_NAME = "ASR_Summary.xlsx"

RENAMED_XBRL_DIR_PATH = XBRL_ROOT_PATH + "/" + RENAMED_XBRL_DIR_NAME
EDINET_INFO_FILE_PATH = XBRL_ROOT_PATH + "/" + EDINET_INFO_FILE_NAME
XBRL_CONTENTS_FILE_PATH = RENAMED_XBRL_DIR_PATH + "/" + XBRL_CONTENTS_FILE_NAME
last_year = 2017 # 最終年度

