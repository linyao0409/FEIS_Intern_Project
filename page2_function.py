

######################################## 匯入套件######################################## import pandas as pd
import openpyxl
import xlrd
import os
from os import listdir, getcwd
from os.path import isfile, join
import sys
import datetime
#import win32com.client as win32
import time
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from email.mime.multipart import MIMEMultipart
import blpapi
from xbbg import blp
import pandas as pd
######################################## Bloomberg連接########################################

def process_data():
    ######################################## Bloomberg連接########################################
    options = blpapi.SessionOptions()
    options.setServerHost('localhost') # setServerHost()方法將連接的Bloomberg API服務器主機設置為localhost。這意味著您正在連接到本地機器上運行的Bloomberg API服務器。如果您使用的是遠程服務器，您需要將該遠程服務器的主機地址設置為此處。
    options.setServerPort(8194)
    session = blpapi.Session(options) # 使用blpapi.Session()建立一個會話(session)物件，並將剛剛創建的options物件作為參數傳遞給它。
    session.start() # 使用start()方法啟動會話。這將嘗試建立到Bloomberg API服務器的連接。如果連接成功，您就可以使用該會話來與Bloomberg API進行數據交互。
    ######################################## 開始製作表格########################################



    ################# load the Isin table#################################
    file_path = "ISIN_code_test.xlsx"  # 記得替換成你的Excel檔案路徑
    df = pd.read_excel(file_path, engine="openpyxl")
    df["Corp"] = df["ISIN"] + " Corp"
    df.set_index(df["Corp"],inplace = True)

    ############長度###########################################
    length = len(df.index)
    isinIndex_series = df["ISIN"]
    isinIndex = list(isinIndex_series)
    ########################################################################
    # R--Corp 
    #R1 = "Corp"
    df["Corp"] = df["ISIN"] + " Corp"
    isinCorp = (df["Corp"]).tolist()
    df.index=isinCorp
    #print(df)

    ################################################################

    col_temp = ["Security_Name","PX_BID","PX_ASK","CBBT_PX_BID","CBBT_PX_ASK","Currency","SETTLE_DT","RTG_SP","RTG_MOODY","RTG_FITCH"]
    aa = blp.bdp(isinCorp,col_temp)
    df = df.join(aa)

    df['px_bid'].fillna(0, inplace=True)
    df['px_bid'].fillna(0, inplace=True)
    df['cbbt_px_bid'].fillna(0, inplace=True)
    df['cbbt_px_ask'].fillna(0, inplace=True)
    ###################################################################



    # BGN_PX_BID
    isin_BGN  = isinIndex_series + " @BGN Corp"
    bgn_b_a = blp.bdp(isin_BGN,["BID","ASK"])
    bgn_b_a.index = bgn_b_a.index.str.split("@").str[0]
    bgn_b_a.index = bgn_b_a.index + "Corp"
    bgn_b_a.rename(columns={"bid":"bgn_px_bid","ask":"bgn_px_ask"},inplace = True)
    df = df.join(bgn_b_a)

    df['bgn_px_bid'].fillna(0, inplace=True)
    df['bgn_px_ask'].fillna(0, inplace=True)


    # BMRK_PX_BID
    t = isinIndex_series + " @BMRK CORP"
    bt = blp.bdp(t,"PX_BID")
    if bt.empty:
        bt.index = isinCorp
        bt["bmrk_px_bid"] = 0
        df = df.join(bt)
    else:
        bt.index = bt.index.str.split("@").str[0]
        bt.index = bt.index + "Corp"
        bt.rename(columns={"px_bid":"bmrk_px_bid"},inplace=True)
        df = df.join(bt)
    #BMRK_PX_ASK
    at = blp.bdp(t,"PX_ASK")
    if at.empty:
        at.index = isinCorp
        at["bmrk_px_ask"] = 0
        df = df.join(at)
    else:
        at.index = at.index.str.split("@").str[0]
        at.index = at.index + "Corp"
        at.rename(columns={"px_ask":"bmrk_px_ask"},inplace =True)
        df = df.join(at)
        

    df['bmrk_px_bid'].fillna(0, inplace=True)
    df['bmrk_px_ask'].fillna(0, inplace=True)

    my = blp.bdp(isinCorp,"MTY_YEARS")
    if my.empty:
        my.index = isinCorp
        my["mty_years"] = 999
        df = df.join(my)
    else:
        df = df.join(my)
        df['mty_years'].fillna(999,inplace=True)
        my["mty_years"] = my["mty_years"].round(2)
        

    #################################################################################################
    # round(2) ???
    columns_to_round = ["px_bid", "px_ask","cbbt_px_bid","cbbt_px_ask","bgn_px_bid","bgn_px_ask","bmrk_px_bid","bmrk_px_ask"]
    df[columns_to_round] = df[columns_to_round].round(2)


    #RTG - sp moody fitch
    #r_smf = blp.bdp(isinCorp,["RTG_SP","RTG_MOODY","RTG_FITCH"])
    #df = df.join(r_smf)
    ###########################################
    file_path = "rating.xlsx"  # 記得替換成你的Excel檔案路徑
    df_rating = pd.read_excel(file_path, engine="openpyxl")
    ###########################################
    z_sp = zip(isinCorp,df["rtg_sp"])
    for i,x in z_sp:
        if pd.isnull(x):
            df.loc[i,"sp"] = 99
        else:
            y = df_rating.loc[df_rating['S&P'] == x,'degree'].iloc[0]
            df.loc[i,"sp"] = y

    z_moody = zip(isinCorp,df["rtg_moody"])
    for i,x in z_moody:
        if pd.isnull(x):
            df.loc[i,"moody"] = 99
        else:
            y = df_rating.loc[df_rating["Moody's"] == x,'degree'].iloc[0]
            df.loc[i,"moody"] = y

    z_fitch = zip(isinCorp,df["rtg_fitch"])
    for i,x in z_fitch:
        if pd.isnull(x):
            df.loc[i,"fitch"] = 99
        else:
            y = df_rating.loc[df_rating["Fitch"] == x,'degree'].iloc[0]
            df.loc[i,"fitch"] = y

    for i in isinCorp:
        if df.loc[i,'sp'] < 10 or df.loc[i,'moody'] < 10 or df.loc[i,'fitch'] < 10:
            df.loc[i,"分類"] = "投等"
        else:
            df.loc[i,"分類"] = "非投等"

    ##################################################################################################################

    # E a buffer for bid
    # F a buffer for ask

    for i in isinCorp:
        temp1 = df.loc[i,"mty_years"]
        temp2 = df.loc[i,'分類']
        ans1 = 0
        ans2 = 0
        if temp1 > 20 and temp2 == '非投等':
            ans1 = -0.25
            ans2 = 0.25
        else:
            if temp1 < 20 and temp2 == '投等':
                ans1 = -0.15
                ans2 = 0.15
            else:
                ans1 = -0.2
                ans2 = 0.2
        df.loc[i,"a_buffer_for_bid"] = ans1
        df.loc[i,"a_buffer_for_ask"] = ans2
    ###################################################################################
    file_path = "exchange_adjust.xlsx"  # Make sure to replace this with your actual Excel file path
    ea_df = pd.read_excel(file_path, engine="openpyxl", index_col="判斷B")
    for i in isinCorp:
        temp = df.loc[i,"currency"]
        temp1 = ea_df.loc[temp,"Bid"]
        temp2 = ea_df.loc[temp,"Ask"]

        df.loc[i,"b_buffer_for_bid"] = temp1
        df.loc[i,"b_buffer_for_ask"] = temp2


    # C PX_BID

    for i in isinCorp:
        temp_currency = df.loc[i,"currency"] #q
        temp_bgn_px_bid = df.loc[i,"bgn_px_bid"] #m
        temp_cbbt_px_bid = df.loc[i,"cbbt_px_bid"] #k
        temp_px_bid = df.loc[i,"px_bid"] #i
        temp_a_buffer_for_bid = df.loc[i,"a_buffer_for_bid"] # e
        temp_b_buffer_for_bid = df.loc[i,"b_buffer_for_bid"] # g
        ans = 0

        if temp_currency == "USD":
            if temp_bgn_px_bid == 0:
                if temp_cbbt_px_bid == 0:
                    ans = temp_px_bid
                else:
                    ans = temp_cbbt_px_bid
            else:
                ans = temp_bgn_px_bid
            ans = ans + (temp_a_buffer_for_bid-temp_b_buffer_for_bid)
        
        else:
            if temp_cbbt_px_bid == 0:
                ans = temp_px_bid
            else:
                ans = temp_cbbt_px_bid
            ans = ans + (temp_a_buffer_for_bid-temp_b_buffer_for_bid)
        df.loc[i,"PX_BID"] = ans

    ###############################################
    # D PX_ASK

    for i in isinCorp:
        temp_currency = df.loc[i,"currency"] #q
        temp_bgn_px_ask = df.loc[i,"bgn_px_ask"] #n
        temp_cbbt_px_ask = df.loc[i,"cbbt_px_ask"] # l
        temp_px_ask = df.loc[i,"px_ask"] # j
        temp_a_buffer_for_ask = df.loc[i,"a_buffer_for_ask"] # f
        temp_b_buffer_for_ask = df.loc[i,"b_buffer_for_ask"] # h
        ans = 0

        if temp_currency == "USD":
            if temp_bgn_px_ask == 0:
                if temp_cbbt_px_ask == 0:
                    ans = temp_px_ask
                else:
                    ans = temp_cbbt_px_ask
            else:
                ans = temp_bgn_px_ask
            ans = ans + (temp_a_buffer_for_ask + temp_b_buffer_for_ask)
        
        else:
            if temp_cbbt_px_ask == 0:
                ans = temp_px_ask
            else:
                ans = temp_cbbt_px_ask
            ans = ans + (temp_a_buffer_for_ask-temp_b_buffer_for_ask)
        df.loc[i,"PX_ASK"] = ans




    #output_file = 'table_815.xlsx' # 覆蓋掉原先ISIN_code_test.clsx檔案
    output_file = f"table_one_{str(datetime.date.today())}"
    #df.to_excel(output_file, index=False)# 將DataFrame輸出成Excel檔案
    print(f"table_one_{str(datetime.date.today())} done")

    ###################################################################
    col_temp = ["CRNCY","MATURITY","NXT_CALL_DT","MIN_PIECE","LQA_LIQUIDITY_SCORE","LEAD_MGR","DS218","DS021"]

    aa = blp.bdp(isinCorp,col_temp)
    df = df.join(aa)
    #aa.close()


    df["ds218"]  = df["ds218"] / 1000
    df["ds021"]  = df["ds021"] / 1000

    df["ds021_ds218_ratio"] = df["ds021"] / df["ds218"]

    ##################################################
    # yld_ytm_ask may have some problem need to be check.
    # 4 para not 2 in bond price excel file
    test = blp.bdp(isinCorp,"YLD_YTM_ASK")
    df = df.join(test)
    df["yld_ytm_ask"] = df["yld_ytm_ask"] / 100

    ##################################################
    # int acc may have some problem need to be check.
    # 4 para not 2 in bond price excel file
    test = blp.bdp(isinCorp,"int acc")
    df = df.join(test)
    df["int_acc"] = df["int_acc"] * 100
    #print(df.iloc[:,-12:])
    ##################################################################
    #讀入昨天檔案取得昨天BID&ASK
    y_df = pd.read_excel("yesterday.xlsx",engine="openpyxl")
    y_df.set_index(y_df.columns[0],inplace=True)
    y_B = y_df["PX_BID"]
    y_B.rename("bid_yesterday",inplace = True)
    y_A = y_df["PX_ASK"]
    y_A.rename("ask_yesterday",inplace = True)

    df = df.join(y_B)
    df = df.join(y_A)

    #print(df.iloc[:,-14:])


    #df.iloc[0, -2] = 78.21

    df["bid_difference"] = df["PX_BID"] - df["bid_yesterday"]
    df["ask_difference"] = df["PX_ASK"] - df["ask_yesterday"]

    # 取絕對值
    df["bid_difference"] = df["bid_difference"].abs()
    df["ask_difference"] = df["ask_difference"].abs()


    output_file = f"table_two_{str(datetime.date.today())}.xlsx"
    df.to_excel(output_file, index=False)# 將DataFrame輸出成Excel
    print(f"{output_file} done")




if __name__ == "__main__":
    process_data()
    