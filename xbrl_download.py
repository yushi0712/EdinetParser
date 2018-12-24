# -*- coding: utf-8 -*-
"""
Created on Wed Dec 12 22:16:55 2018

@author: e12135
"""

import os
import pandas as pd

from os import path
from edinet_xbrl.edinet_xbrl_downloader import EdinetXbrlDownloader


EdinetCodeFilePath = "C:/local/users/nishimura/Documents/!Theme/財務分析/EDINETコード.xlsx"
XBRL_DIR_PATH = "C:/Users/e12135/OneDrive - Konica Minolta/Archives/HDJA-UT/財務データ/XBRL"
#XBRL_DIR_PATH = "C:/local/users/nishimura/Documents/!Theme/財務分析/XBRL"

# EDINETコードExcelの内容をDataFrameにインポート
EdinetCodeFile = pd.ExcelFile(EdinetCodeFilePath)
sheet_names = EdinetCodeFile.sheet_names
df_sheet = EdinetCodeFile.parse(sheet_names[0], skiprows=[0])

# XBRLのダウンロード対象となる提出者のリストを作成しDataFrameに格納する
count = 0
dir_list = list()
column_list = ["EDINET Code", "Name", "Dir"]
df_presenter = pd.DataFrame(columns=column_list)
for index, row in df_sheet.iterrows():
    if row["Flag1"] == 1:
        presenter = pd.DataFrame(columns=column_list)
        presenter.loc[0] = [row["ＥＤＩＮＥＴコード"], row["提出者名"], row["ＥＤＩＮＥＴコード"]+"_"+row["提出者名"]]
        df_presenter = df_presenter.append(presenter)
        count = count + 1
print("Count:", count)

# 提出者ごとのフォルダ作成
count = 0
total = len(df_presenter)
xbrl_downloader = EdinetXbrlDownloader()
for index, row in df_presenter.iterrows():
    # フォルダ作成
    target_dir = XBRL_DIR_PATH + "/" + row["Dir"]
    count = count + 1
    if not os.path.isdir(target_dir):
        os.mkdir(target_dir)
    # ダウンロード    
    ticker = row["EDINET Code"]
    xbrl_downloader.download_by_ticker(ticker, target_dir)

    print(str(count)+"/"+str(total), row["EDINET Code"], row["Name"])
    
            
            
'''
from edinet_xbrl.edinet_xbrl_downloader import EdinetXbrlDownloader

## init downloader
xbrl_downloader = EdinetXbrlDownloader()

## set a ticker you want to download xbrl file
ticker = "E00989"
target_dir = "C:/local/users/nishimura/Programs/!Python/sandbox/Edinet/XBRL"
xbrl_downloader.download_by_ticker(ticker, target_dir)
'''