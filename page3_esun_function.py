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


def esun_table():
    file_path = f"table_two_{str(datetime.date.today())}.xlsx"
    esun_df = pd.read_excel(file_path,engine="openpyxl")
    esun_df = esun_df[esun_df["ESUN"] == 1]
    #print(type(esun_df['maturity'][0]))
    #print(esun_df.columns)

    # Credit Rating 
    esun_df["Credit Rating"] = esun_df["rtg_sp"] + "/" + esun_df["rtg_moody"]
    #print(esun_df["Credit Rating"])

    # Maturity 
    """ already exist"""

    # Tensor to Call
    # !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    esun_df["Tensor to Call"] = 0
    #print(esun_df["Tensor to Call"])

    # Tensor 
    today = datetime.datetime.combine(datetime.date.today(), datetime.datetime.min.time())  # 将日期部分与虚拟时间部分组合
    esun_df["Tensor"] = (esun_df["maturity"] - today).dt.days
    esun_df["Tensor"] = esun_df["Tensor"] / 365
    #print(esun_df["Tensor"])
    #print(esun_df.iloc[:,-10:])

    # index - > isin_corp
    isin_corp = esun_df["Corp"]
    esun_df.set_index(["Corp"],inplace=True)

    #CCY already exist

    # S/D already exist  ( settle_dt)

    # Clean Bid already exist(PX_BID)
    # Clean Ask already exist(PX_ASK)

    # Accured Interest already exist(int_acc)
    esun_df["int_acc"] = esun_df["int_acc"] / 100

    #Dirty Bid
    #Dirty Ask
    esun_df["Dirty Bid"] = esun_df["PX_BID"] + esun_df["int_acc"]
    esun_df["Dirty Ask"] = esun_df["PX_ASK"] + esun_df["int_acc"]

    # UF
    # !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    esun_df["UF"] = 1


    # Clean Ask+ UF
    esun_df["Clean Ask+ UF"] = esun_df["PX_ASK"] + esun_df["UF"]

    # Clean YTC Clean YTM & After Fee YTC & After Fee YTM
    esun_df["Clean YTC"] = 0
    esun_df["Clean YTM"] = 0
    esun_df["After Fee YTC"] = 0
    esun_df["After Fee YTM"] = 0
    esun_df["市場券源狀況"] = ""
    #print(esun_df.iloc[:,-15:])


    esun_df.rename(columns={"security_name":"Name","maturity":"Maturity","crncy":"CCY","settle_dt":"S/D","PX_BID":"Clean Bid","PX_ASK":"Clean Ask","int_acc":"Accrued Interest"},inplace=True)

    selected_cols = ['Name', 'Credit Rating', 'Maturity', 'Tensor to Call', 'Tensor', 'ISIN', 'CCY', 'S/D', 'Clean Bid', 'Clean Ask', 'Accrued Interest', 'Dirty Bid', 'Dirty Ask', 'UF', 'Clean Ask+ UF', 'Clean YTC', 'Clean YTM', 'After Fee YTC', 'After Fee YTM', '市場券源狀況']
    new_df = esun_df[selected_cols].copy()

    #print(new_df)

    output_file = f"ESUN_table_{str(datetime.date.today())}.xlsx"
    new_df.to_excel(output_file, index=False)# 將DataFrame輸出成Excel
    print(f"{output_file} done")







if __name__ == "__main__":
    esun_table()

    