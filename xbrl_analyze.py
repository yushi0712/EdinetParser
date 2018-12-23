# xbrl から読み込む
import os
import sys
import pandas as pd
import xbrl_common
import time

sys.path.append(os.path.join(os.path.dirname(__file__), 'XbrlReader'))

from os import path
from xbrl_proc import read_xbrl

ASR_SUMMARY_FILE_NAME = "ASR_Summary.xlsx"


def _get_tag_val(df, tags_and_contexts):
    val = ""
    for tc in tags_and_contexts:
        # tagを探す
        df_tag = df[df["tag"] == tc[0]]
        l = len(list(df_tag[df_tag == True].index))
        if l == 0: # tagが見つからなかった
            continue

        # contextを探す
        df_context = df_tag[df_tag["context"] == tc[1]]
        l = len(list(df_context[df_context == True].index))
        if l == 0: # contextが見つからなかった
            continue

        val = df_context["値"].values[0]

    return val



#---------------------------------------
# XBRL Contentsファイルを読み込む
#---------------------------------------
xbrl_contents_file = pd.ExcelFile(xbrl_common.XBRL_CONTENTS_FILE_PATH)
df_xbrl_contents = xbrl_contents_file.parse(xbrl_contents_file.sheet_names[0], skipcols=[0])

#---------------------------------------
# XBRLファイルを読み込む -> ASR_Summaary作成
#---------------------------------------
summary_column = ["EDINETコード", "提出者名", "証券コード", "業種", "年度", "会計基準", "従業員数", "総資産", "売上高", "純利益", "営業CF", "投資CF", "財務CF", "売上／人員"]
df_asr_summary = pd.DataFrame(columns=summary_column)
for index, row in df_xbrl_contents.iterrows():
    start_time = time.perf_counter()
    # XBRLファイルのパス生成
    xbrl_path = xbrl_common.RENAMED_XBRL_DIR_PATH + r"/" + row["オリジナルファイル"]
    if path.isfile(xbrl_path):
        if row["年度"] == 2018:
            # 基本情報
            s_asr = pd.Series(index=summary_column)
            s_asr["EDINETコード"] = row["EDINETコード"]
            s_asr["提出者名"] = row["提出者名"]
            s_asr["証券コード"] = row["証券コード"]
            s_asr["業種"] = row["業種"]
            s_asr["年度"] = row["年度"]
            # 財務情報
            df_xbrl_data = read_xbrl(xbrl_path)
            CYI = "CurrentYearInstant"
            CYD = "CurrentYearDuration"
            CYD_NCM = "CurrentYearDuration_NonConsolidatedMember"
            s_asr["会計基準"] = _get_tag_val(df_xbrl_data, [["AccountingStandardsDEI", "FilingDateInstant"]])
            s_asr["従業員数"] = _get_tag_val(df_xbrl_data, [["NumberOfEmployeeIFRS",CYI], ["NumberOfEmployees",CYI]])
            s_asr["総資産"] = _get_tag_val(df_xbrl_data, [["TotalAssetsIFRSSummaryOfBusinessResults",CYI], ["TotalAssetsSummaryOfBusinessResults",CYI], ["TotalAssetsUSGAAPSummaryOfBusinessResults",CYI]])
            s_asr["売上高"] = _get_tag_val(df_xbrl_data, [["RevenueIFRSSummaryOfBusinessResults",CYD], ["NetSalesSummaryOfBusinessResults",CYD], ["RevenuesUSGAAPSummaryOfBusinessResults",CYD]])
            s_asr["純利益"] = _get_tag_val(df_xbrl_data, [["ProfitLossAttributableToOwnersOfParentIFRSSummaryOfBusinessResults",CYD], ["ProfitLossAttributableToOwnersOfParentSummaryOfBusinessResults",CYD], ["NetIncomeLossAttributableToOwnersOfParentUSGAAPSummaryOfBusinessResults",CYD]])
            s_asr["営業CF"] = _get_tag_val(df_xbrl_data, [["CashFlowsFromUsedInOperatingActivitiesIFRSSummaryOfBusinessResults",CYD], ["NetCashProvidedByUsedInOperatingActivitiesSummaryOfBusinessResults",CYD], ["CashFlowsFromUsedInOperatingActivitiesUSGAAPSummaryOfBusinessResults",CYD]])
            s_asr["投資CF"] = _get_tag_val(df_xbrl_data, [["CashFlowsFromUsedInInvestingActivitiesIFRSSummaryOfBusinessResults",CYD], ["NetCashProvidedByUsedInInvestingActivitiesSummaryOfBusinessResults",CYD], ["CashFlowsFromUsedInInvestingActivitiesUSGAAPSummaryOfBusinessResults",CYD]])
            s_asr["財務CF"] = _get_tag_val(df_xbrl_data, [["CashFlowsFromUsedInFinancingActivitiesIFRSSummaryOfBusinessResults",CYD], ["NetCashProvidedByUsedInFinancingActivitiesSummaryOfBusinessResults",CYD], ["CashFlowsFromUsedInFinancingActivitiesUSGAAPSummaryOfBusinessResults",CYD]])
            s_asr["売上／人員"] = s_asr["売上高"] / s_asr["従業員数"]
            # DataFrameに追加          
            df_asr_summary = df_asr_summary.append(s_asr, ignore_index=True)
            # 経過表示
            print("{0:.1f}[sec]".format(time.perf_counter()-start_time), index+1, "/", len(df_xbrl_contents),\
                  row["EDINETコード"], row["提出者名"], row["年度"])
    # Debug用break
    #break;
#---------------------------------------
# ASR_SummaaryをExcelファイルに保存
#---------------------------------------
df_asr_summary.to_excel(xbrl_common.XBRL_ROOT_PATH + "/" + ASR_SUMMARY_FILE_NAME)

