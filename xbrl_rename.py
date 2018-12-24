# xbrl から読み込む
import os
import re
import sys
import shutil
import pandas as pd
import xbrl_common

sys.path.append(os.path.join(os.path.dirname(__file__), 'XbrlReader'))

from os import path


##########[CLASS]##############
class XbrlFile:
    def __init__(self, dir_name, file_name, year):
        self.dir_name = dir_name
        self.file_name = file_name
        self.year = year
    
##########[CLASS]##############
class XbrlPresenter:
    def __init__(self, name, edinet_code, securities_code, industry, xbrl_files):
        self.name = name
        self.edinet_code = edinet_code
        self.securities_code = securities_code
        self.industry = industry
        self.xbrl_files = xbrl_files


last_year = xbrl_common.last_year + 1

#---------------------------------------------------------
# EDINETの情報をExcelファイルから取得しEDINET情報DataFrameを作成
#---------------------------------------------------------
edinet_info_file = pd.ExcelFile(xbrl_common.EDINET_INFO_FILE_PATH)
df_edinet_info = edinet_info_file.parse(edinet_info_file.sheet_names[0], skiprows=[0])

#---------------------------------------------------------
# EDINET情報DataFrameから業種と証券コードを取り出す
#---------------------------------------------------------
industry_dict = dict()
securities_code_dict = dict()
for index, row in df_edinet_info.iterrows():
    tmp_code = row["ＥＤＩＮＥＴコード"]
    industry_dict[tmp_code] = row["提出者業種"]
    securities_code_dict[tmp_code] = str(row["証券コード"])
    

#--------------------------------------------------------
# XBRL関連ファイルの中から必要な(-asr-)ファイルのみを抽出しリスト化
#---------------------------------------------------------
# フォルダ取得
all_items = os.listdir(xbrl_common.ORIGINAL_XBRL_DIR_PATH)
all_dirs = [f for f in all_items if os.path.isdir(os.path.join(xbrl_common.ORIGINAL_XBRL_DIR_PATH, f))]

presenter_list = list()
for _dir in all_dirs: # 各フォルダ
    edinet_code = _dir[0:6]
    presenter_name = _dir[7:]
    xbrl_dir_path = xbrl_common.ORIGINAL_XBRL_DIR_PATH + "/" + _dir
    if path.isdir(xbrl_dir_path):
        # 初期化
        target_dict = dict()
        for y in range(last_year, last_year-10, -1): # 各年度のファイルを抽出
             target_dict[str(y)]=0   
        # 各年の最新の有価証券報告書ファイルを選択する
        files = os.listdir(xbrl_dir_path)
        asr_files = [f for f in files if "-asr-" in f] # -asr-:有価証券報告書(-asr-)のみを抽出
        target_asr_file = dict()
        for asr in asr_files:
            tmp = re.split("[-_]", asr)
            if int(tmp[8]) > target_dict[tmp[5]]:
                target_dict[tmp[5]] = int(tmp[8]) # tmp5が年度 tmp[8]が通番
                target_asr_file[tmp[5]] = asr
        # target_filesに各年の最新rev.の有価証券報告書ファイルを格納する      
        xbrl_files = list()
        for y in range(last_year, last_year-10, -1): # 各年度のファイルを抽出
            str_y = str(y)
            if target_dict[str_y] != 0: 
                inst_xbrl_file = XbrlFile(_dir, target_asr_file[str_y], str_y)
                xbrl_files.append(inst_xbrl_file)
            
        inst_presenter = XbrlPresenter(presenter_name, edinet_code,\
                                       securities_code=securities_code_dict[edinet_code],\
                                       industry=industry_dict[edinet_code],\
                                       xbrl_files=xbrl_files)
        presenter_list.append(inst_presenter)

#---------------------------------------------------------
# XBRL_ContentsのDataFrameを作成
#---------------------------------------------------------
print("◆XBRL_ContentsのDataFrameを作成")
contents_column = ["EDINETコード", "提出者名", "証券コード", "業種", "年度", "フォルダ", "オリジナルファイル", "リネームファイル"]
df_xbrl_contetns = pd.DataFrame(columns=contents_column)
for index, p in enumerate(presenter_list):
    for xbrl_file in p.xbrl_files:
        new_file_name = p.edinet_code+"_"+p.name+"_"+xbrl_file.year+"_"+p.industry+".xbrl"
        df = pd.Series([p.edinet_code, p.name, str(p.securities_code),\
                        p.industry, xbrl_file.year, xbrl_file.dir_name, xbrl_file.file_name, new_file_name], index=contents_column)
        df_xbrl_contetns = df_xbrl_contetns.append(df, ignore_index=True)

    if index%100==0:
        print("\r{0}/{1}".format(index, len(presenter_list)), end="")

print("  -> 完了")

#---------------------------------------------------------
# ファイルをRename
#---------------------------------------------------------
print("◆ファイルをRename")
if not path.isdir(xbrl_common.RENAMED_XBRL_DIR_PATH): # フォルダが無いときは作成する
    os.mkdir(xbrl_common.RENAMED_XBRL_DIR_PATH)
# Contentsファイル（Excel）を作成
df_xbrl_contetns.to_excel(xbrl_common.XBRL_CONTENTS_FILE_PATH)
# ファイルをリネームしてコピー
for index, row in df_xbrl_contetns.iterrows():
    org_path = xbrl_common.ORIGINAL_XBRL_DIR_PATH + "/" + row["フォルダ"] + "/" + row["オリジナルファイル"]
    dest_path = xbrl_common.RENAMED_XBRL_DIR_PATH + "/" + row["リネームファイル"]
    # Temporary:XBRLのParserがファイル名で解析しているので一旦オリジナルファイルと同じ名前にする
    #dest_path = xbrl_common.RENAMED_XBRL_DIR_PATH + "/" + row["オリジナルファイル"]
    if path.isfile(org_path):
        shutil.copyfile(org_path, dest_path)

    if index%100==0:
        print("\r{0}/{1}".format(index, len(df_xbrl_contetns)), end="")

print("  -> 完了")
                   
#print(xbrl_files) 

