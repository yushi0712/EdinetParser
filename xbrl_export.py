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
        break

    return val


#---------------------------------------
# XBRL Contentsファイルを読み込む
#---------------------------------------
print("◆XBRL Contentsファイルを読み込む", end="")
xbrl_contents_file = pd.ExcelFile(xbrl_common.XBRL_CONTENTS_FILE_PATH)
df_xbrl_contents = xbrl_contents_file.parse(xbrl_contents_file.sheet_names[0], skipcols=[0])
print("  -> 完了")

#---------------------------------------
# XBRLファイルを読み込み、ASR_Summaary作成
#---------------------------------------
print("◆XBRLファイルを読み込み、ASR_Summaary作成")
summary_column = ["EDINETコード", "提出者名", "証券コード", "業種", "年度", "会計基準",\
                  "従業員数", "総資産", "売上高", "純利益", "株価収益率", "営業CF", "投資CF", "財務CF", "現金",\
                  "従業員数(P1Y)", "総資産(P1Y)", "売上高(P1Y)", "純利益(P1Y)", "株価収益率(P1Y)", "営業CF(P1Y)", "投資CF(P1Y)", "財務CF(P1Y)", "現金(P1Y)",\
                  "従業員数(P2Y)", "総資産(P2Y)", "売上高(P2Y)", "純利益(P2Y)", "株価収益率(P2Y)", "営業CF(P2Y)", "投資CF(P2Y)", "財務CF(P2Y)", "現金(P2Y)",\
                  "従業員数(P3Y)", "総資産(P3Y)", "売上高(P3Y)", "純利益(P3Y)", "株価収益率(P3Y)", "営業CF(P3Y)", "投資CF(P3Y)", "財務CF(P3Y)", "現金(P3Y)",\
                  "従業員数(P4Y)", "総資産(P4Y)", "売上高(P4Y)", "純利益(P4Y)", "株価収益率(P4Y)", "営業CF(P4Y)", "投資CF(P4Y)", "財務CF(P4Y)", "現金(P4Y)"]
df_asr_summary = pd.DataFrame(columns=summary_column)
# 短縮形
CYI = "CurrentYearInstant"
P1I = "Prior1YearInstant"
P2I = "Prior2YearInstant"
P3I = "Prior3YearInstant"
P4I = "Prior4YearInstant"
CYI_NCM = "CurrentYearInstant_NonConsolidatedMember"
P1I_NCM = "Prior1YearInstant_NonConsolidatedMember"
P2I_NCM = "Prior2YearInstant_NonConsolidatedMember"
P3I_NCM = "Prior3YearInstant_NonConsolidatedMember"
P4I_NCM = "Prior4YearInstant_NonConsolidatedMember"
CYD = "CurrentYearDuration"
P1D = "Prior1YearDuration"
P2D = "Prior2YearDuration"
P3D = "Prior3YearDuration"
P4D = "Prior4YearDuration"
CYD_NCM = "CurrentYearDuration_NonConsolidatedMember"
P1D_NCM = "Prior1YearDuration_NonConsolidatedMember"
P2D_NCM = "Prior2YearDuration_NonConsolidatedMember"
P3D_NCM = "Prior3YearDuration_NonConsolidatedMember"
P4D_NCM = "Prior4YearDuration_NonConsolidatedMember"
industry_list = ["電気機器", "精密機器", "機械", "化学", "情報・通信業", "輸送用機器", "非鉄金属", "医薬品", "鉄鋼"]
start_time = time.perf_counter()
for index, row in df_xbrl_contents.iterrows():
    # XBRLファイルのパス生成
    xbrl_path = xbrl_common.RENAMED_XBRL_DIR_PATH + r"/" + row["リネームファイル"]
    if path.isfile(xbrl_path):
        keyword = row["業種"]
        if (row["年度"]==2017):
            # 基本情報
            s_asr = pd.Series(index=summary_column)
            s_asr["EDINETコード"] = row["EDINETコード"]
            s_asr["提出者名"] = row["提出者名"]
            s_asr["証券コード"] = row["証券コード"]
            s_asr["業種"] = row["業種"]
            s_asr["年度"] = row["年度"]
            # 財務情報
            df_xbrl_data = read_xbrl(xbrl_path, row["オリジナルファイル"])
            s_asr["会計基準"] = _get_tag_val(df_xbrl_data, [["AccountingStandardsDEI", "FilingDateInstant"]])
            
            s_asr["従業員数"] = _get_tag_val(df_xbrl_data, [["NumberOfEmployeesIFRSSummaryOfBusinessResults",CYI], ["NumberOfEmployeeIFRS",CYI], ["NumberOfEmployees",CYI], ["NumberOfEmployees",CYI_NCM]])
            s_asr["総資産"] = _get_tag_val(df_xbrl_data, [["TotalAssetsIFRSSummaryOfBusinessResults",CYI], ["TotalAssetsSummaryOfBusinessResults",CYI], ["TotalAssetsUSGAAPSummaryOfBusinessResults",CYI], ["TotalAssetsSummaryOfBusinessResults",CYI_NCM]])
            s_asr["売上高"] = _get_tag_val(df_xbrl_data, [["RevenueIFRSSummaryOfBusinessResults",CYD], ["NetSalesSummaryOfBusinessResults",CYD], ["RevenuesUSGAAPSummaryOfBusinessResults",CYD], ["NetSalesSummaryOfBusinessResults",CYD_NCM], ["OperatingRevenue1SummaryOfBusinessResults",CYD], ["OrdinaryIncomeSummaryOfBusinessResults",CYD], ["OperatingRevenue2SummaryOfBusinessResults",CYD], ["WholeChainStoreSalesSummaryOfBusinessResults",CYD], ["NetSalesAndOperatingRevenue2SummaryOfBusinessResults",CYD], ["NetSalesOfCompletedConstructionContractsSummaryOfBusinessResults",CYD]])
            s_asr["純利益"] = _get_tag_val(df_xbrl_data, [["ProfitLossAttributableToOwnersOfParentIFRSSummaryOfBusinessResults",CYD], ["ProfitLossAttributableToOwnersOfParentSummaryOfBusinessResults",CYD], ["NetIncomeLossAttributableToOwnersOfParentUSGAAPSummaryOfBusinessResults",CYD], ["NetIncomeLossSummaryOfBusinessResults",CYD_NCM]])
            s_asr["株価収益率"] = _get_tag_val(df_xbrl_data, [["PriceEarningsRatioIFRSSummaryOfBusinessResults",CYD], ["PriceEarningsRatioSummaryOfBusinessResults",CYD], ["PriceEarningsRatioUSGAAPSummaryOfBusinessResults",CYD], ["PriceEarningsRatioSummaryOfBusinessResults",CYD_NCM]])
            s_asr["営業CF"] = _get_tag_val(df_xbrl_data, [["CashFlowsFromUsedInOperatingActivitiesIFRSSummaryOfBusinessResults",CYD], ["NetCashProvidedByUsedInOperatingActivitiesSummaryOfBusinessResults",CYD], ["CashFlowsFromUsedInOperatingActivitiesUSGAAPSummaryOfBusinessResults",CYD], ["NetCashProvidedByUsedInOperatingActivitiesSummaryOfBusinessResults",CYD_NCM]])
            s_asr["投資CF"] = _get_tag_val(df_xbrl_data, [["CashFlowsFromUsedInInvestingActivitiesIFRSSummaryOfBusinessResults",CYD], ["NetCashProvidedByUsedInInvestingActivitiesSummaryOfBusinessResults",CYD], ["CashFlowsFromUsedInInvestingActivitiesUSGAAPSummaryOfBusinessResults",CYD], ["NetCashProvidedByUsedInInvestingActivitiesSummaryOfBusinessResults",CYD_NCM]])
            s_asr["財務CF"] = _get_tag_val(df_xbrl_data, [["CashFlowsFromUsedInFinancingActivitiesIFRSSummaryOfBusinessResults",CYD], ["NetCashProvidedByUsedInFinancingActivitiesSummaryOfBusinessResults",CYD], ["CashFlowsFromUsedInFinancingActivitiesUSGAAPSummaryOfBusinessResults",CYD], ["NetCashProvidedByUsedInFinancingActivitiesSummaryOfBusinessResults",CYD_NCM]])
            s_asr["現金"] = _get_tag_val(df_xbrl_data, [["CashAndCashEquivalentsIFRSSummaryOfBusinessResults",CYI], ["CashAndCashEquivalentsSummaryOfBusinessResults",CYI], ["CashAndCashEquivalentsUSGAAPSummaryOfBusinessResults",CYI], ["CashAndCashEquivalentsSummaryOfBusinessResults",CYI_NCM]])

            s_asr["従業員数(P1Y)"] = _get_tag_val(df_xbrl_data, [["NumberOfEmployeesIFRSSummaryOfBusinessResults",P1I], ["NumberOfEmployeeIFRS",P1I], ["NumberOfEmployees",P1I], ["NumberOfEmployees",P1I_NCM]])
            s_asr["総資産(P1Y)"] = _get_tag_val(df_xbrl_data, [["TotalAssetsIFRSSummaryOfBusinessResults",P1I], ["TotalAssetsSummaryOfBusinessResults",P1I], ["TotalAssetsUSGAAPSummaryOfBusinessResults",P1I], ["TotalAssetsSummaryOfBusinessResults",P1I_NCM]])
            s_asr["売上高(P1Y)"] = _get_tag_val(df_xbrl_data, [["RevenueIFRSSummaryOfBusinessResults",P1D], ["NetSalesSummaryOfBusinessResults",P1D], ["RevenuesUSGAAPSummaryOfBusinessResults",P1D], ["NetSalesSummaryOfBusinessResults",P1D_NCM], ["OperatingRevenue1SummaryOfBusinessResults",P1D], ["OrdinaryIncomeSummaryOfBusinessResults",P1D], ["OperatingRevenue2SummaryOfBusinessResults",P1D], ["WholeChainStoreSalesSummaryOfBusinessResults",P1D], ["NetSalesAndOperatingRevenue2SummaryOfBusinessResults",P1D], ["NetSalesOfCompletedConstructionContractsSummaryOfBusinessResults",P1D]])
            s_asr["純利益(P1Y)"] = _get_tag_val(df_xbrl_data, [["ProfitLossAttributableToOwnersOfParentIFRSSummaryOfBusinessResults",P1D], ["ProfitLossAttributableToOwnersOfParentSummaryOfBusinessResults",P1D], ["NetIncomeLossAttributableToOwnersOfParentUSGAAPSummaryOfBusinessResults",P1D], ["NetIncomeLossSummaryOfBusinessResults",P1D_NCM]])
            s_asr["株価収益率(P1Y)"] = _get_tag_val(df_xbrl_data, [["PriceEarningsRatioIFRSSummaryOfBusinessResults",P1D], ["PriceEarningsRatioSummaryOfBusinessResults",P1D], ["PriceEarningsRatioUSGAAPSummaryOfBusinessResults",P1D], ["PriceEarningsRatioSummaryOfBusinessResults",P1D_NCM]])
            s_asr["営業CF(P1Y)"] = _get_tag_val(df_xbrl_data, [["CashFlowsFromUsedInOperatingActivitiesIFRSSummaryOfBusinessResults",P1D], ["NetCashProvidedByUsedInOperatingActivitiesSummaryOfBusinessResults",P1D], ["CashFlowsFromUsedInOperatingActivitiesUSGAAPSummaryOfBusinessResults",P1D], ["NetCashProvidedByUsedInOperatingActivitiesSummaryOfBusinessResults",P1D_NCM]])
            s_asr["投資CF(P1Y)"] = _get_tag_val(df_xbrl_data, [["CashFlowsFromUsedInInvestingActivitiesIFRSSummaryOfBusinessResults",P1D], ["NetCashProvidedByUsedInInvestingActivitiesSummaryOfBusinessResults",P1D], ["CashFlowsFromUsedInInvestingActivitiesUSGAAPSummaryOfBusinessResults",P1D], ["NetCashProvidedByUsedInInvestingActivitiesSummaryOfBusinessResults",P1D_NCM]])
            s_asr["財務CF(P1Y)"] = _get_tag_val(df_xbrl_data, [["CashFlowsFromUsedInFinancingActivitiesIFRSSummaryOfBusinessResults",P1D], ["NetCashProvidedByUsedInFinancingActivitiesSummaryOfBusinessResults",P1D], ["CashFlowsFromUsedInFinancingActivitiesUSGAAPSummaryOfBusinessResults",P1D], ["NetCashProvidedByUsedInFinancingActivitiesSummaryOfBusinessResults",P1D_NCM]])
            s_asr["現金(P1Y)"] = _get_tag_val(df_xbrl_data, [["CashAndCashEquivalentsIFRSSummaryOfBusinessResults",P1I], ["CashAndCashEquivalentsSummaryOfBusinessResults",P1I], ["CashAndCashEquivalentsUSGAAPSummaryOfBusinessResults",P1I], ["CashAndCashEquivalentsSummaryOfBusinessResults",P1I_NCM]])

            s_asr["従業員数(P2Y)"] = _get_tag_val(df_xbrl_data, [["NumberOfEmployeesIFRSSummaryOfBusinessResults",P1I], ["NumberOfEmployeeIFRS",P2I], ["NumberOfEmployees",P2I], ["NumberOfEmployees",P2I_NCM]])
            s_asr["総資産(P2Y)"] = _get_tag_val(df_xbrl_data, [["TotalAssetsIFRSSummaryOfBusinessResults",P2I], ["TotalAssetsSummaryOfBusinessResults",P2I], ["TotalAssetsUSGAAPSummaryOfBusinessResults",P2I], ["TotalAssetsSummaryOfBusinessResults",P2I_NCM]])
            s_asr["売上高(P2Y)"] = _get_tag_val(df_xbrl_data, [["RevenueIFRSSummaryOfBusinessResults",P2D], ["NetSalesSummaryOfBusinessResults",P2D], ["RevenuesUSGAAPSummaryOfBusinessResults",P2D], ["NetSalesSummaryOfBusinessResults",P2D_NCM], ["OperatingRevenue1SummaryOfBusinessResults",P2D], ["OrdinaryIncomeSummaryOfBusinessResults",P2D], ["OperatingRevenue2SummaryOfBusinessResults",P2D], ["WholeChainStoreSalesSummaryOfBusinessResults",P2D], ["NetSalesAndOperatingRevenue2SummaryOfBusinessResults",P2D], ["NetSalesOfCompletedConstructionContractsSummaryOfBusinessResults",P2D]])
            s_asr["純利益(P2Y)"] = _get_tag_val(df_xbrl_data, [["ProfitLossAttributableToOwnersOfParentIFRSSummaryOfBusinessResults",P2D], ["ProfitLossAttributableToOwnersOfParentSummaryOfBusinessResults",P2D], ["NetIncomeLossAttributableToOwnersOfParentUSGAAPSummaryOfBusinessResults",P2D], ["NetIncomeLossSummaryOfBusinessResults",P2D_NCM]])
            s_asr["株価収益率(P2Y)"] = _get_tag_val(df_xbrl_data, [["PriceEarningsRatioIFRSSummaryOfBusinessResults",P2D], ["PriceEarningsRatioSummaryOfBusinessResults",P2D], ["PriceEarningsRatioUSGAAPSummaryOfBusinessResults",P2D], ["PriceEarningsRatioSummaryOfBusinessResults",P2D_NCM]])
            s_asr["営業CF(P2Y)"] = _get_tag_val(df_xbrl_data, [["CashFlowsFromUsedInOperatingActivitiesIFRSSummaryOfBusinessResults",P2D], ["NetCashProvidedByUsedInOperatingActivitiesSummaryOfBusinessResults",P2D], ["CashFlowsFromUsedInOperatingActivitiesUSGAAPSummaryOfBusinessResults",P2D], ["NetCashProvidedByUsedInOperatingActivitiesSummaryOfBusinessResults",P2D_NCM]])
            s_asr["投資CF(P2Y)"] = _get_tag_val(df_xbrl_data, [["CashFlowsFromUsedInInvestingActivitiesIFRSSummaryOfBusinessResults",P2D], ["NetCashProvidedByUsedInInvestingActivitiesSummaryOfBusinessResults",P2D], ["CashFlowsFromUsedInInvestingActivitiesUSGAAPSummaryOfBusinessResults",P2D], ["NetCashProvidedByUsedInInvestingActivitiesSummaryOfBusinessResults",P2D_NCM]])
            s_asr["財務CF(P2Y)"] = _get_tag_val(df_xbrl_data, [["CashFlowsFromUsedInFinancingActivitiesIFRSSummaryOfBusinessResults",P2D], ["NetCashProvidedByUsedInFinancingActivitiesSummaryOfBusinessResults",P2D], ["CashFlowsFromUsedInFinancingActivitiesUSGAAPSummaryOfBusinessResults",P2D], ["NetCashProvidedByUsedInFinancingActivitiesSummaryOfBusinessResults",P2D_NCM]])
            s_asr["現金(P2Y)"] = _get_tag_val(df_xbrl_data, [["CashAndCashEquivalentsIFRSSummaryOfBusinessResults",P2I], ["CashAndCashEquivalentsSummaryOfBusinessResults",P2I], ["CashAndCashEquivalentsUSGAAPSummaryOfBusinessResults",P2I], ["CashAndCashEquivalentsSummaryOfBusinessResults",P2I_NCM]])

            s_asr["従業員数(P3Y)"] = _get_tag_val(df_xbrl_data, [["NumberOfEmployeesIFRSSummaryOfBusinessResults",P1I], ["NumberOfEmployeeIFRS",P3I], ["NumberOfEmployees",P3I], ["NumberOfEmployees",P3I_NCM]])
            s_asr["総資産(P3Y)"] = _get_tag_val(df_xbrl_data, [["TotalAssetsIFRSSummaryOfBusinessResults",P3I], ["TotalAssetsSummaryOfBusinessResults",P3I], ["TotalAssetsUSGAAPSummaryOfBusinessResults",P3I], ["TotalAssetsSummaryOfBusinessResults",P3I_NCM]])
            s_asr["売上高(P3Y)"] = _get_tag_val(df_xbrl_data, [["RevenueIFRSSummaryOfBusinessResults",P3D], ["NetSalesSummaryOfBusinessResults",P3D], ["RevenuesUSGAAPSummaryOfBusinessResults",P3D], ["NetSalesSummaryOfBusinessResults",P3D_NCM], ["OperatingRevenue1SummaryOfBusinessResults",P3D], ["OrdinaryIncomeSummaryOfBusinessResults",P3D], ["OperatingRevenue2SummaryOfBusinessResults",P3D], ["WholeChainStoreSalesSummaryOfBusinessResults",P3D], ["NetSalesAndOperatingRevenue2SummaryOfBusinessResults",P3D], ["NetSalesOfCompletedConstructionContractsSummaryOfBusinessResults",P3D]])
            s_asr["純利益(P3Y)"] = _get_tag_val(df_xbrl_data, [["ProfitLossAttributableToOwnersOfParentIFRSSummaryOfBusinessResults",P3D], ["ProfitLossAttributableToOwnersOfParentSummaryOfBusinessResults",P3D], ["NetIncomeLossAttributableToOwnersOfParentUSGAAPSummaryOfBusinessResults",P3D], ["NetIncomeLossSummaryOfBusinessResults",P3D_NCM]])
            s_asr["株価収益率(P3Y)"] = _get_tag_val(df_xbrl_data, [["PriceEarningsRatioIFRSSummaryOfBusinessResults",P3D], ["PriceEarningsRatioSummaryOfBusinessResults",P3D], ["PriceEarningsRatioUSGAAPSummaryOfBusinessResults",P3D], ["PriceEarningsRatioSummaryOfBusinessResults",P3D_NCM]])
            s_asr["営業CF(P3Y)"] = _get_tag_val(df_xbrl_data, [["CashFlowsFromUsedInOperatingActivitiesIFRSSummaryOfBusinessResults",P3D], ["NetCashProvidedByUsedInOperatingActivitiesSummaryOfBusinessResults",P3D], ["CashFlowsFromUsedInOperatingActivitiesUSGAAPSummaryOfBusinessResults",P3D], ["NetCashProvidedByUsedInOperatingActivitiesSummaryOfBusinessResults",P3D_NCM]])
            s_asr["投資CF(P3Y)"] = _get_tag_val(df_xbrl_data, [["CashFlowsFromUsedInInvestingActivitiesIFRSSummaryOfBusinessResults",P3D], ["NetCashProvidedByUsedInInvestingActivitiesSummaryOfBusinessResults",P3D], ["CashFlowsFromUsedInInvestingActivitiesUSGAAPSummaryOfBusinessResults",P3D], ["NetCashProvidedByUsedInInvestingActivitiesSummaryOfBusinessResults",P3D_NCM]])
            s_asr["財務CF(P3Y)"] = _get_tag_val(df_xbrl_data, [["CashFlowsFromUsedInFinancingActivitiesIFRSSummaryOfBusinessResults",P3D], ["NetCashProvidedByUsedInFinancingActivitiesSummaryOfBusinessResults",P3D], ["CashFlowsFromUsedInFinancingActivitiesUSGAAPSummaryOfBusinessResults",P3D], ["NetCashProvidedByUsedInFinancingActivitiesSummaryOfBusinessResults",P3D_NCM]])
            s_asr["現金(P3Y)"] = _get_tag_val(df_xbrl_data, [["CashAndCashEquivalentsIFRSSummaryOfBusinessResults",P3I], ["CashAndCashEquivalentsSummaryOfBusinessResults",P3I], ["CashAndCashEquivalentsUSGAAPSummaryOfBusinessResults",P3I], ["CashAndCashEquivalentsSummaryOfBusinessResults",P3I_NCM]])

            s_asr["従業員数(P4Y)"] = _get_tag_val(df_xbrl_data, [["NumberOfEmployeesIFRSSummaryOfBusinessResults",P1I], ["NumberOfEmployeeIFRS",P4I], ["NumberOfEmployees",P4I], ["NumberOfEmployees",P4I_NCM]])
            s_asr["総資産(P4Y)"] = _get_tag_val(df_xbrl_data, [["TotalAssetsIFRSSummaryOfBusinessResults",P4I], ["TotalAssetsSummaryOfBusinessResults",P4I], ["TotalAssetsUSGAAPSummaryOfBusinessResults",P4I], ["TotalAssetsSummaryOfBusinessResults",P4I_NCM]])
            s_asr["売上高(P4Y)"] = _get_tag_val(df_xbrl_data, [["RevenueIFRSSummaryOfBusinessResults",P4D], ["NetSalesSummaryOfBusinessResults",P4D], ["RevenuesUSGAAPSummaryOfBusinessResults",P4D], ["NetSalesSummaryOfBusinessResults",P4D_NCM], ["OperatingRevenue1SummaryOfBusinessResults",P4D], ["OrdinaryIncomeSummaryOfBusinessResults",P4D], ["OperatingRevenue2SummaryOfBusinessResults",P4D], ["WholeChainStoreSalesSummaryOfBusinessResults",P4D], ["NetSalesAndOperatingRevenue2SummaryOfBusinessResults",P4D], ["NetSalesOfCompletedConstructionContractsSummaryOfBusinessResults",P4D]])
            s_asr["純利益(P4Y)"] = _get_tag_val(df_xbrl_data, [["ProfitLossAttributableToOwnersOfParentIFRSSummaryOfBusinessResults",P4D], ["ProfitLossAttributableToOwnersOfParentSummaryOfBusinessResults",P4D], ["NetIncomeLossAttributableToOwnersOfParentUSGAAPSummaryOfBusinessResults",P4D], ["NetIncomeLossSummaryOfBusinessResults",P4D_NCM]])
            s_asr["株価収益率(P4Y)"] = _get_tag_val(df_xbrl_data, [["PriceEarningsRatioIFRSSummaryOfBusinessResults",P4D], ["PriceEarningsRatioSummaryOfBusinessResults",P4D], ["PriceEarningsRatioUSGAAPSummaryOfBusinessResults",P4D], ["PriceEarningsRatioSummaryOfBusinessResults",P4D_NCM]])
            s_asr["営業CF(P4Y)"] = _get_tag_val(df_xbrl_data, [["CashFlowsFromUsedInOperatingActivitiesIFRSSummaryOfBusinessResults",P4D], ["NetCashProvidedByUsedInOperatingActivitiesSummaryOfBusinessResults",P4D], ["CashFlowsFromUsedInOperatingActivitiesUSGAAPSummaryOfBusinessResults",P4D], ["NetCashProvidedByUsedInOperatingActivitiesSummaryOfBusinessResults",P4D_NCM]])
            s_asr["投資CF(P4Y)"] = _get_tag_val(df_xbrl_data, [["CashFlowsFromUsedInInvestingActivitiesIFRSSummaryOfBusinessResults",P4D], ["NetCashProvidedByUsedInInvestingActivitiesSummaryOfBusinessResults",P4D], ["CashFlowsFromUsedInInvestingActivitiesUSGAAPSummaryOfBusinessResults",P4D], ["NetCashProvidedByUsedInInvestingActivitiesSummaryOfBusinessResults",P4D_NCM]])
            s_asr["財務CF(P4Y)"] = _get_tag_val(df_xbrl_data, [["CashFlowsFromUsedInFinancingActivitiesIFRSSummaryOfBusinessResults",P4D], ["NetCashProvidedByUsedInFinancingActivitiesSummaryOfBusinessResults",P4D], ["CashFlowsFromUsedInFinancingActivitiesUSGAAPSummaryOfBusinessResults",P4D], ["NetCashProvidedByUsedInFinancingActivitiesSummaryOfBusinessResults",P4D_NCM]])
            s_asr["現金(P4Y)"] = _get_tag_val(df_xbrl_data, [["CashAndCashEquivalentsIFRSSummaryOfBusinessResults",P4I], ["CashAndCashEquivalentsSummaryOfBusinessResults",P4I], ["CashAndCashEquivalentsUSGAAPSummaryOfBusinessResults",P4I], ["CashAndCashEquivalentsSummaryOfBusinessResults",P4I_NCM]])

            # DataFrameに追加          
            df_asr_summary = df_asr_summary.append(s_asr, ignore_index=True)
            # 経過表示
            #print("{0:.1f}[sec]".format(time.perf_counter()-start_time), index+1, "/", len(df_xbrl_contents),\
            #      row["EDINETコード"], row["提出者名"], row["年度"])
    print("\r{0}/{1} ({2})".format(index, len(df_xbrl_contents), len(df_asr_summary)), end="")
    # Debug用break
    if len(df_asr_summary) > 100:
        sys.exit()
    #break;
elapsed_time = time.perf_counter() - start_time
print("  -> 完了", "平均処理時間:{0:.3f}秒".format(elapsed_time/len(df_asr_summary)))

#---------------------------------------
# ASR_SummaaryをExcelファイルに保存
#---------------------------------------
print("◆ASR_SummaaryをExcelファイルに保存", end="")
df_asr_summary.to_excel(xbrl_common.XBRL_ROOT_PATH + "/" + ASR_SUMMARY_FILE_NAME)
print("  -> 完了")

