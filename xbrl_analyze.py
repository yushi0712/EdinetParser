import os
import sys
import pandas as pd
import xbrl_common
import time

from os import path
from xbrl_proc import read_xbrl


#---------------------------------------
# ASR Summaryファイルを読み込む
#---------------------------------------
print("◆XBRL Contentsファイルを読み込む", end="")
asr_summary_file = pd.ExcelFile(xbrl_common.XBRL_ROOT_PATH + "/" + xbrl_common.ASR_SUMMARY_FILE_NAME)
df_asr_summary = asr_summary_file.parse(sheet_name="OrgData")
df_industry = pd.DataFrame(list(set(df_asr_summary["業種"])), columns=["業種"])
print("  -> 完了")

# 財務CFと株価収益率以外の空欄をカウント
df_asr_summary["空欄数"]=-1
df_asr_summary["テスト"]=-1
for index, row in df_asr_summary.iterrows():
    null_num = row[6:].isnull().sum()
    cnt = 0
    for idx, itm in row.iteritems():
        if ("財務CF" in idx) or ("株価収益率" in idx):
            if pd.isnull(itm):
                cnt += 1
    df_asr_summary.at[index, "空欄数"] = null_num - cnt

#---------------------------------------
# 業種ごとにDataFrameを作成
#---------------------------------------
df_specified_industry = dict()    
for index, row in df_industry.iterrows():
    q_word = "業種 == \"{0}\"".format(row["業種"])
    df_specified_industry[row["業種"]] = df_asr_summary.query(q_word)


