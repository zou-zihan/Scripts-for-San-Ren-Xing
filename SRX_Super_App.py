#! /usr/bin/env python3

import numpy as np
import pandas as pd
import datetime as dt
import re
import math
import unicodedata
import os
import sys
import requests
import pygsheets
import string
import pickle
from cryptography.fernet import Fernet
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from tqdm import tqdm
import pyfiglet
import time
import uuid
import ast

'''
generate pdf works well with borb version 2.1.5.2
pip uninstall borb
pip install --upgrade borb==2.1.5.2
'''
from borb.pdf.document.document import Document as borb_Document
from borb.pdf.page.page import Page as borb_Page
from decimal import Decimal
from borb.pdf.canvas.layout.page_layout.multi_column_layout import MultiColumnLayout as borb_MCL
from borb.pdf.canvas.layout.page_layout.multi_column_layout import SingleColumnLayout as borb_SCL
from borb.pdf.canvas.layout.page_layout.page_layout import PageLayout as borb_PageLayout
from borb.pdf.canvas.layout.text.paragraph import Paragraph as borb_Paragraph
from borb.pdf.pdf import PDF as borb_PDF
import pathlib
from borb.pdf.canvas.layout.image.image import Image as borb_Image
from decimal import Decimal
from borb.pdf.canvas.layout.table.fixed_column_width_table import (
    FixedColumnWidthTable as borb_Table,
)
from borb.pdf.canvas.layout.table.flexible_column_width_table import (
    FlexibleColumnWidthTable as borb_flex_Table,
)
from borb.pdf.canvas.layout.layout_element import Alignment as borb_align
from borb.pdf.canvas.font.simple_font.true_type_font import TrueTypeFont as borb_TrueTypeFont
from borb.pdf.canvas.layout.table.table import TableCell as borb_TableCell
from borb.pdf import HexColor as borb_HexColor

'''
import requests
import urllib
import os
from cryptography.fernet import Fernet

def on_internet():
    url = "https://www.google.com.sg/"
    timeout = 5
    try:
        request = requests.get(url, timeout=timeout)
        return True
    except (requests.ConnectionError, requests.Timeout) as exception:
        return False

githubUserName = "zou-zihan"
githubRepoName = "Scripts-for-San-Ren-Xing"
githubBranchName = "main"
script_backup_filename = "Encrypted Script.txt"
backup_foldername = "NABK"

#打包盒管理负责人(加引号加逗号)
box_on_duty = ""

#酒水管理负责人(加引号加逗号)
drink_on_duty = ""

#排班负责人(加引号)
payslip_on_duty = ""

on_net = on_internet()

if not on_net:
    wifi = False

    print("警告! 没有网络连接! ")
    print("离线模式下，程序将会全程采用本地备份数据来运行")
    print("可能有一些没有及时更新的资料而导致生成的文件不准确")
    print("非常不推荐你使用离线模式！")
    print()
    for i, name in enumerate(["离线模式继续运行","终止运行"]):
        print("{}扣{}".format(name, i))

    try:
        userInputOne = int(input("在这里输入>>>: "))
    except:
        userInputOne = 1
        print("无效命令")

    if userInputOne == 0:
        with open("fernet_key.txt", "rb") as keyfile:
            fernet_key = keyfile.read()

        with open("{}/{}/{}".format(os.getcwd(), backup_foldername, script_backup_filename), "rb") as sfile:
            script = sfile.read()

        fernet_handler = Fernet(fernet_key)
        script = fernet_handler.decrypt(script).decode()
        exec(script)

    elif userInputOne == 1:
        pass

    else:
        print("无效命令")

else:
    wifi = True
    URL = "{}/{}/{}/{}/{}".format("https://raw.githubusercontent.com", githubUserName, githubRepoName, githubBranchName, "SRX_Super_App.py")
    script = urllib.request.urlopen(URL).read().decode()
    exec(script)

#中午营业额
lun_sales = 0.0

#中午顾客人数
lun_gc = 0

#中午楼面员工人数(加引号)
lun_fwc = ""

#中午厨房员工人数(加引号)
lun_kwc = ""

#下午茶楼面员工人数(加引号)
tb_fwc = ""

#下午茶厨房员工人数(加引号)
tb_kwc = ""

#下午茶营业额
tb_sales = 0.0

#下午茶顾客人数
tb_gc = ""

#晚上楼面员工人数(加引号)
night_fwc = ""

#晚上厨房员工人数(加引号)
night_kwc = ""

#收银员(加引号加逗号)
cashier_on_duty = ""

#经理(加引号加逗号)
manager_on_duty = ""

#不显示主菜单(True/False)
do_not_show_menu = True

exec(open("Local Script.py", encoding="utf-8").read())

'''

database_url = "gAAAAABksmmj1NC7ooHFZ7wTi9d1dhk-Q6i1Y_GWWb24xVpLfC0lmNa5vsXGDhCYwRO_NLy8NDge8xV6zna6ny1pDagdp77f8eE9nrOmRWdzeEYMi21hbEEY88G0lnfJPl0XZb2kN6M7BWasrkIp3TeDgSFmdnigSvlT4NJn92lVdCKJIJXCq5Qik4qj52Q-JJU7LZ5Xc_Lf"
db_setting_url = "gAAAAABksml2tGtCXmmE6x14Fj-_f9RKeL7pMvBlrDn1puSotmBKVAZzWRaYLQq0FRhe2MHYX6JQMDzzeUKI5lKjADOYSKWMaA3EKicdBC94U6KDqjMf71prmgXaVKu-dq_W9NBJtgDKKU7HfmRyKU6mE1Di1Nge1wt4FFuHRjJqIxvMO6eX9OQHOY8mEDYDkcLfP0GKvXu2"

rtn_database_url = "gAAAAABlPkHLRchmBxC7TlvW5-b6okgDr9SbK3UcLFXUMTvnnYExtq1ad3zAW8LVhZi5FygUynb-s84mgvYQocJhF0cFS_w5iF4unfJqMvusbVta7pyLzTT3C5qFpz3deMYrV1i0UWLAOMQnuGX6J8zTKYaDv3S5k1VnOuLzu3IoUO2GhxBFIg5Gku6zk1TlJsRZVC2ghPi9"
rtn_control_url = "gAAAAABlPkHkSIrx4kubK9zovDQhCUjv9Wa4b9QWQ98ss2fyetRVrlQXJvw2ITBk-TqPL5K8knWKnEs84RyQ4RoBThRmqgEC_FkZ25aEfdje0KwQJUdmyfTZWoYDyARJsiMSz4xzrQoitURokO-zuEzrnRM1AuPZeFZrY0jukzQtbu-RDWbwuPApXCeESbkHtydubmx98lW5"

service_filename = "google auth bot.json"
serialized_rule_filename = "Serialized Rules.pkl"
constants_sheetname = "Constants"
box_num = 40
drink_num = 12
promo_num = 30

google_auth = pygsheets.authorize(service_file=service_filename)

def get_key():
    with open("fernet_key.txt", "rb") as f:
        fernet_key = f.read()

    fernet_obj = Fernet(fernet_key)
    test_byte = b'gAAAAABksWqIeorgYZDLzrO4c7v1bLwXOMzqGqCoAQhzZvnYJvsmf8tgsW0H1yAZbTo-D8DmudoyXp3HJ2KuiQDP7VWJZtV9ow=='

    try:
        test_result = fernet_obj.decrypt(test_byte)
        return fernet_key

    except:
        fernet_key = 0
        return fernet_key

def get_outlet():
    for filename in os.listdir():
        if filename.endswith("_book.xlsx"):
            outlet = filename.replace("_book.xlsx", "")

    return outlet.strip().upper()

def parseGoogleHTMLSheet(df):

    df.drop("Unnamed: 0", axis=1, inplace=True)
    df.columns = df.iloc[0,:]
    df.drop(0, axis=0,inplace=True)

    col_1st = df.columns[0]
    df.drop(df[df[col_1st].isnull()].index,axis=0, inplace=True)

    df.reset_index(inplace=True)
    df.drop("index", axis=1, inplace=True)

    return df

def remove_spaces(var):
    return str(var).translate({ord(y): None for y in string.whitespace})

def normal_round(n, decimals=0):
    expoN = n * 10 ** decimals
    if abs(expoN) - abs(math.floor(expoN)) < 0.5:
        return math.floor(expoN) / 10 ** decimals
    return math.ceil(expoN) / 10 ** decimals

def single_zero(x):
    if float(x) == 0:
        return int(0)
    else:
        return x

def integer_check(x):
    try:
        x = int(x)
        return True
    except ValueError:
        return False

def fernet_decrypt(encrypted_string, fernet_key):
    fernet_obj = Fernet(fernet_key)

    return fernet_obj.decrypt(encrypted_string.encode()).decode()

def need_integer(x, absolute=False):
    if integer_check(x) and absolute:
        return abs(int(x))

    elif integer_check(x) and not absolute:
        return int(x)

def option_num(options_list):
    options_list = np.array(options_list).astype(str)
    rx = np.arange(len(options_list)).astype(str)
    options_list_kou_int = []
    for i in range(len(options_list)):
        options_list_kou_int += [options_list[i]+"扣"+rx[i]]
    
    options_list_kou_int = np.array(options_list_kou_int)

    for item in options_list_kou_int:
        print(item)
    
    return options_list_kou_int
    
def option_limit(options_list, user_input):
    if integer_check(user_input):
        user_input = int(user_input)
        if user_input in range(len(options_list)):
            return user_input
        else:
            while user_input not in range(len(options_list)):
                print('输入的'+ str(user_input)+ '无效!')
                user_input = need_integer(input('请重新输入: '), absolute=True)
                while user_input in range(len(options_list)):
                    return user_input
    else:
        while user_input not in range(len(options_list)):
                print('输入的'+ str(user_input)+ '无效!')
                user_input = need_integer(input('请重新输入: '), absolute=True)
                while user_input in range(len(options_list)):
                    return user_input

def get_dfs(google_auth, db_setting_url, constants_sheetname, serialized_rule_filename, outlet, fernet_key, wifi, backup_foldername):

    if wifi:
        try:
            db_setting_url = fernet_decrypt(db_setting_url, fernet_key)

            db_setting_sheet = google_auth.open_by_url(db_setting_url)

            #k_df
            k_sheet_index = db_setting_sheet.worksheet(property="title", value=constants_sheetname).index
            k_df = db_setting_sheet[k_sheet_index].get_as_df()

            outlet_col_index = k_df.columns.get_loc(key=str(outlet).strip().upper())

            k_dict = {}
            for index in range(len(k_df)):
                k_dict.update({ str(k_df.iloc[index, 0]) : str(k_df.iloc[index, outlet_col_index]) })

            #value_dict
            vd_sheet_name = str(k_dict["value_dict_rule_sheetname"])
            vd_url = fernet_decrypt(str(k_dict["value_dict_url"]), fernet_key).strip()

            if vd_url == db_setting_url:
                vd_sheet_index = db_setting_sheet.worksheet(property="title", value=vd_sheet_name).index
                vd_df = db_setting_sheet[vd_sheet_index].get_as_df()

            else:
                vd_sheet = google_auth.open_by_url(vd_url)
                vd_sheet_index = vd_sheet.worksheet(property="title", value=vd_sheet_name).index
                vd_df = vd_sheet[vd_sheet_index].get_as_df()

            #takeaway_box_rule
            tb_sheet_name = str(k_dict["takeaway_box_rule_sheetname"])
            tb_url = fernet_decrypt(str(k_dict["takeaway_box_rule_url"]), fernet_key).strip()

            if tb_url == db_setting_url:
                tb_sheet_index = db_setting_sheet.worksheet(property="title", value=tb_sheet_name).index
                tb_df = db_setting_sheet[tb_sheet_index].get_as_df()

            else:
                tb_sheet = google_auth.open_by_url(tb_url)
                tb_sheet_index = tb_sheet.worksheet(property="title",value=tb_sheet_name).index
                tb_df = tb_sheet[tb_sheet_index].get_as_df()

            #extra_takeaway_box_rule
            etb_sheet_name = str(k_dict["extra_takeaway_box_rule_sheetname"])
            etb_url = fernet_decrypt(str(k_dict["extra_takeaway_box_rule_url"]), fernet_key).strip()

            if etb_url == db_setting_url:
                etb_sheet_index = db_setting_sheet.worksheet(property="title", value = etb_sheet_name).index
                etb_df = db_setting_sheet[etb_sheet_index].get_as_df()

            else:
                etb_sheet = google_auth.open_by_url(etb_url)
                etb_sheet_index = etb_sheet.worksheet(property="title", value = etb_sheet_name).index
                etb_df = etb_sheet[etb_sheet_index].get_as_df()

            #print_rule
            pr_sheet_name = str(k_dict["print_rule_sheetname"])
            pr_url = fernet_decrypt(str(k_dict["print_rule_url"]), fernet_key).strip()

            if pr_url == db_setting_url:
                pr_sheet_index = db_setting_sheet.worksheet(property="title", value = pr_sheet_name).index
                pr_df = db_setting_sheet[pr_sheet_index].get_as_df()

            else:
                pr_sheet = google_auth.open_by_url(pr_url)
                pr_sheet_index = pr_sheet.worksheet(property="title", value = pr_sheet_name).index
                pr_df = pr_sheet[pr_sheet_index].get_as_df()


            rule_df_dict = {"vd_df" : vd_df,
                            "tb_df" : tb_df,
                            "etb_df" : etb_df,
                            "pr_df" : pr_df,
                           }

            backup_list = [k_df, rule_df_dict]
            with open("{}/{}/{}".format(os.getcwd(), backup_foldername, serialized_rule_filename), "wb") as f:
                pickle.dump(backup_list, f)

        except Exception as e:
            print()
            print("从网络读取规则文件失败，错误描述如下: ")
            print()
            print(e)

            with open("{}/{}/{}".format(os.getcwd(), backup_foldername, serialized_rule_filename), "rb") as f:
                backup_list = pickle.load(f)

            k_df = backup_list[0]
            rule_df_dict = backup_list[1]

            outlet_col_index = k_df.columns.get_loc(key=outlet)

            k_dict = {}
            for index in range(len(k_df)):
                k_dict.update({ str(k_df.iloc[index, 0]) : str(k_df.iloc[index, outlet_col_index]) })

            print("已采用本地备份规则文件数据")

    else:
        with open("{}/{}/{}".format(os.getcwd(), backup_foldername, serialized_rule_filename), "rb") as f:
            backup_list = pickle.load(f) #a list of two objects, 0 is k_df, 1 is rule_df_dict

        k_df = backup_list[0]
        rule_df_dict = backup_list[1]

        outlet_col_index = k_df.columns.get_loc(key=outlet)

        k_dict = {}
        for index in range(len(k_df)):
            k_dict.update({ str(k_df.iloc[index, 0]) : str(k_df.iloc[index, outlet_col_index]) })

    return k_dict, rule_df_dict

def get_book_dfs(k_dict, outlet):

    book_filename = str(k_dict["book_filename"])
    min_book_len_allowable = int(k_dict["min_book_len_allowable"])

    read = pd.read_excel(book_filename, thousands=",")
    columnsRename = np.arange(0, len(read.columns)).astype(str)
    read.columns = columnsRename
    read["0"] = read["0"].astype(str)

    if len(read) < int(min_book_len_allowable):
        print("你应该没有复制所有的后台数据, 复制完所有的数据后再运行一遍")
        is_local_book = None
        read = None
        pay_breakdown = None
        pay_brkdwn = None
        read_ = None

    else:
        is_local_book = not read[read["0"].str.contains("GST Reg No")].empty

        if is_local_book:
            try:
                outlet_loc = outlet.strip().capitalize()
            except NameError:
                outlet_loc = "门店【待编辑】"

            print("The program had detected that you are using a book from local machine to do night audit")
            print()
            print("Important messages to take note:")
            print("In order to copy all contents from HTML correctly you have to use a computer for copy-and-paste action")
            print("Open the HTML file in a browser, press CTRL A and CTRL C on Windows/Linux")
            print("CMD A and CMD C on Mac for copy")
            print()
            print("DO NOT use mouse to drag and select to copy, this programme will raise errors if you do that.")
            print()
            print("Create a new excel file, and CTRL V directly on Windows/Linux for paste")
            print("If you did pasting on a Mac, you have to Select 'Paste Special' -> 'Paste Special...'")
            print("'Source:' is 'paste', 'As:' is 'Unicode Text', then click 'ok' to paste the texts into an excel file.")
            print()
            print()
            print("Please remove detail payment breakdown.")
            print("It is between the last payment method and immediately before '**OPENING CASH BALANCE'.")
            print()

            try:
                tsbt_index = int(read[read["0"] == 'Total Sales Before Tax & Srv Chg'].index[0])

            except IndexError:
                tsbt_index = int(read[read["0"] == "Total Sales Before Tax"].index[0])

            read.iloc[tsbt_index, 0] = 'Total Sales Before Tax & Srv Charge'

            gst_index = int(read[read["0"] == "GST 9%"].index[0])
            read.iloc[int(gst_index), 0] = "GST"

            svc_index = int(read[read["0"] == "SCV CHG 10%"].index[0])
            read.iloc[int(svc_index), 0] = "Service Charge 10%"

            gstReg_index = int(read[read["0"].str.contains("GST Reg No")].index[0])
            read.iloc[int(gstReg_index), 0] = "Goods and Services Tax Reg No"

            continue_ = True

        else:
            try:
                bi = int(read[read['0'].str.contains('Date:')].index.values[0])
                ei = int(read[read['0'] == 'Total Sales'].index.values[0])
                outlet_loc = str(read.iloc[(bi+1):(ei),:].dropna(axis=1).values[0][0])
                continue_ = True
            except IndexError:
                continue_ = False

        if not continue_:
            print("导出的报表无法识别来自哪家门店")
            print("请选择门店导出报表后再运行一遍")

            read = None
            pay_breakdown = None
            pay_brkdwn = None
            read_ = None

        else:
            pbi = int(read[read['0'].str.contains('PAYMENT BREAKDOWN:')].index[0])

            if is_local_book:
                pei = int(read[read['0'].str.contains('OPENING CASH BALANCE:')].index[0])

            else:
                try:
                    pei = int(read[read['0'].str.contains('PAYMENT BREAKDOWN (POS):', regex=False)].index[0])
                except IndexError:
                    pei = int(read[read['0'].str.contains('Transaction Void Items')].index[0])

            pay_breakdown = read.copy()
            pay_breakdown = pay_breakdown.iloc[(pbi+1):(pei), :]
            pay_breakdown["0"] = pay_breakdown["0"].astype(str)

            drop_index_criteria = (pay_breakdown["0"].str.contains("nan"))|(pay_breakdown["0"]==" ")
            drop_index = pay_breakdown[drop_index_criteria].index

            pay_breakdown.drop(drop_index, axis=0, inplace=True)
            pay_breakdown.reset_index(inplace=True)
            pay_breakdown.drop('index', axis=1, inplace=True)

            dropColumns = []
            for columnIndex in range(len(pay_breakdown.columns)):
                if pay_breakdown[str(columnIndex)].isnull().all():
                    dropColumns += [str(columnIndex)]
                else:
                    continue

            pay_breakdown.drop(dropColumns, axis=1, inplace=True)

            pay_breakdown_reset_columns = np.arange(0, len(pay_breakdown.columns)).astype(str)
            pay_breakdown.columns = pay_breakdown_reset_columns

            pay_brkdwn = pay_breakdown.copy()

            if pay_breakdown.empty:
                print("无任何付款信息")

            else:
                pay_brkdwn["0"] = pay_brkdwn["0"].apply(lambda x:x[:x.find("(")])

            read_ = read.copy()
            read_["0"] = read_["0"].apply(lambda a : unicodedata.normalize("NFKD", a))
            read_["0"] = read_["0"].apply(lambda x : remove_spaces(x))

    return_criteria = (isinstance(is_local_book, bool)) and (isinstance(read, pd.DataFrame))

    if return_criteria:
        book_dict = {
            "is_local_book" : is_local_book,
            "outlet_loc" : outlet_loc,
            "read" : read,
            "pay_breakdown" : pay_breakdown,
            "pay_brkdwn" : pay_brkdwn,
            "read_" : read_
        }

    else:
        book_dict = {}

    return book_dict

def get_df_date(book_dict):

    is_local_book = book_dict["is_local_book"]
    read = book_dict["read"]

    today_date = dt.datetime.today()

    #raw date from book
    if is_local_book:
        try:
            read_date = read.iloc[int(read[read["0"].str.contains("X/Shift Report")].index[0])+1,0]
        except IndexError:
            read_date = read.iloc[int(read[read["0"].str.contains("Z/Closing Report")].index[0])+1,0]

        rdfb = read_date.split()

    else:
        rdfb = read[read["0"].str.contains('Date:')].dropna(axis=1).values[0][0]

    #raw_date_from_book
    if is_local_book:

        try:
            read_date = read.iloc[int(read[read["0"].str.contains("X/Shift Report")].index[0])+1,0]
        except IndexError:
            read_date = read.iloc[int(read[read["0"].str.contains("Z/Closing Report")].index[0])+1,0]

        rdfb = read_date.split()

    else:
        rdfb = read[read["0"].str.contains("Date:")].dropna(axis=1).values[0][0]
        rdfb = rdfb.split()

    if is_local_book:
        dfb = "{} {} {}".format(rdfb[1], rdfb[2], rdfb[3])
        dfb = dt.datetime.strptime(dfb, "%d %b %Y")

    else:
        if len(rdfb) == 8:
            if (rdfb[1] == rdfb[5]) and (rdfb[2] == rdfb[6]) and (rdfb[3] == rdfb[7]):
                #date from book
                dfb = "{} {} {}".format(rdfb[1],rdfb[2],rdfb[3])
                dfb = dt.datetime.strptime(dfb, "%d %b %Y")

            else:
                print("不支持跨日期。 \n报表日期将采用今天日期。")
                dfb = today_date

        else:
            print("日期格式不符，报表将采用今天日期。")
            dfb = today_date

    yesterday = dfb - dt.timedelta(days=1)

    if sys.platform.strip().upper() in ["IOS", "WIN32"]:
        date_chinese_string = dfb.strftime("%YNIAN%mYUE%dRI")
        date_chinese_string = date_chinese_string.replace("NIAN", "年")
        date_chinese_string = date_chinese_string.replace("YUE", "月")
        date_chinese_string = date_chinese_string.replace("RI", "日")

    else:
        date_chinese_string = dfb.strftime("%Y年%m月%d日")

    dfb = pd.to_datetime(dfb)
    yesterday = pd.to_datetime(yesterday)

    date_dict = {
        "dfb" : dfb,
        "yesterday" : yesterday,
        "date_chinese_string" : date_chinese_string
    }

    return date_dict

def get_db_columns(k_dict, drink_num, box_num, promo_num):

    drink_col = ["日期"]
    for n in range(drink_num):
        drink_col += ["{}上日存货".format(k_dict["drink{}".format(n)]),
                      "{}进".format(k_dict["drink{}".format(n)]),
                      "{}出".format(k_dict["drink{}".format(n)]),
                      "{}实结存".format(k_dict["drink{}".format(n)]),
                      "{}备注".format(k_dict["drink{}".format(n)]),
                     ]

    drink_col += ["DATE", "TIME LOG"]

    tabox_col = ["TIME LOG", "DATE"]
    for m in range(box_num):
        tabox_col += ["{}入库".format(k_dict["box{}".format(m)]),
                      "{}出库".format(k_dict["box{}".format(m)]),
                      "{}现有数量".format(k_dict["box{}".format(m)]),
                     ]

    financial_col = ["TIME LOG",
                     "DATE",
                     "DAY",
                     "NET SALES",
                     "CUM NET SALES YESTERDAY",
                     "CUM NET SALES TODAY",
                     "AVG NET SALES PER DAY",
                     "CASH AMOUNT",
                     "DINE IN TOTAL",
                     "TAKEAWAY TOTAL"
                    ]

    promo_col = ["TIME LOG", "DATE"]

    for i in range(promo_num):
        promo_col += ["promo{}_ytd".format(i+1),
                      "promo{}_tdy".format(i+1),
                     ]


    get_columns = {
        "drink_col" : drink_col,
        "tabox_col" : tabox_col,
        "financial_col" : financial_col,
        "promo_col" : promo_col
    }

    return get_columns

def retrieve_database(k_dict, google_auth, fernet_key, wifi, database_url, backup_foldername):

    local_db_filename = k_dict["local_database_filename"]

    drink_db_sheetname = k_dict["drink_database_sheetname"]
    takeaway_box_db_sheetname = k_dict["takeaway_box_database_sheetname"]
    financial_db_sheetname = k_dict["financial_database_sheetname"]
    promotion_db_sheetname = k_dict["promotion_database_sheetname"]

    #get online databases and local databases
    online_databases = {}
    local_databases = {}

    sheet_names = [drink_db_sheetname, takeaway_box_db_sheetname,
                   financial_db_sheetname, promotion_db_sheetname]

    if wifi:
        try:
            online_db_url = fernet_decrypt(database_url, fernet_key).strip()
            online_db_sheet = google_auth.open_by_url(online_db_url)

            for name in sheet_names:
                db_index = online_db_sheet.worksheet(property="title", value = name).index
                db_df = online_db_sheet[db_index].get_as_df()
                online_databases.update({ "online_{}".format(name) : db_df })

                if os.path.exists("{}/{}/{}".format(os.getcwd(), backup_foldername, local_db_filename)):
                    local_db_df = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_db_filename), sheet_name=name)
                    local_databases.update({ "local_{}".format(name) : local_db_df })

            if len(local_databases) == 0:
                take_databases = {}
                for key in online_databases:
                    take_databases.update({ "{}".format(key.replace("online_", "")) : online_databases[key] })

            else:
                take_databases = {}
                for name in sheet_names:
                    open_online_df = online_databases["online_{}".format(name)]
                    open_local_df = local_databases["local_{}".format(name)]

                    open_online_df["DATE"] = pd.to_datetime(open_online_df["DATE"])
                    open_local_df["DATE"] = pd.to_datetime(open_local_df["DATE"])

                    online_max_date_record = open_online_df["DATE"].max()
                    local_max_date_record = open_local_df["DATE"].max()

                    if local_max_date_record > online_max_date_record:
                        take_databases.update({ name : open_local_df })
                    else:
                        take_databases.update({ name : open_online_df })
        except Exception as e:
            print()
            print("从网络读取数据库失败, 错误描述如下: ")
            print()
            print(e)
            take_databases = {}
            for name in sheet_names:
                local_db_df = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_db_filename), sheet_name=name)
                take_databases.update({ name : local_db_df })

            print("已采用本地备份的数据库信息")

    else:
        take_databases = {}
        for name in sheet_names:
            local_db_df = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_db_filename), sheet_name=name)
            take_databases.update({ name : local_db_df })

    return take_databases

def update_columns(take_databases, get_columns, k_dict):

    drink_db_sheetname = k_dict["drink_database_sheetname"]
    takeaway_box_db_sheetname = k_dict["takeaway_box_database_sheetname"]
    financial_db_sheetname = k_dict["financial_database_sheetname"]
    promotion_db_sheetname = k_dict["promotion_database_sheetname"]

    sheet_names = [drink_db_sheetname, takeaway_box_db_sheetname,
                   financial_db_sheetname, promotion_db_sheetname]



    col_dict = {
                sheet_names[0] : get_columns["drink_col"],
                sheet_names[1] : get_columns["tabox_col"],
                sheet_names[2] : get_columns["financial_col"],
                sheet_names[3] : get_columns["promo_col"],
               }

    for key in take_databases:
        df = take_databases[key]
        current_col = df.columns.tolist()

        correct_col = col_dict[key]

        if len(current_col) == len(correct_col):
            if current_col != correct_col:
                df.columns = correct_col
                take_databases[key] = df
                print("{}'s column values had been updated automatically with accordance to the Rule.".format(key))
            else:
                continue
        else:
            print("{}'s column length aren't the same with the Rule.".format(key))

    return take_databases

def box_use_sums(read_, keyname, model_listing, tb_df, k_dict):
    df = read_.copy()

    is_in_criteria = df["0"].isin(model_listing[str(keyname)])

    df = df[is_in_criteria].dropna(axis=1)

    if df.empty:
        return float(0)

    else:
        df = df.iloc[:, :2]

        tb_copy = tb_df.copy()
        tb_copy["MODEL TO USE"] = tb_copy["MODEL TO USE"].apply(lambda x : k_dict[str(x)])

        model_criteria = tb_copy["MODEL TO USE"] == str(keyname)
        tb_copy = tb_copy[model_criteria]
        tb_copy.drop(["CHINESE FOOD NAME", "MODEL TO USE", "ACTIVE STATUS"], axis=1, inplace=True)
        tb_copy.columns = ["0", "MULTIPLIER"]

        merged_df = pd.merge(left=df, right=tb_copy, how="inner")

        merged_df["1"] = merged_df["1"].astype(float)
        merged_df["MULTIPLIER"] = merged_df["MULTIPLIER"].astype(float)

        merged_df["AFTER MULTIPLICATION"] = merged_df["1"] * merged_df["MULTIPLIER"]

        return float(merged_df["AFTER MULTIPLICATION"].sum())

def box_now(tb_data, box_name, box_value, yesterday):

    df = tb_data.copy()
    df["DATE"] = pd.to_datetime(df["DATE"])
    yesterday  = pd.to_datetime(yesterday)

    if float(box_value["{}_god_hand".format(box_name)]) < 0:
        equal_ytd_criteria = df["DATE"] == yesterday
        xianYouShuLiang = df[equal_ytd_criteria]["{}现有数量".format(box_name)].values[0]
        xianYouShuLiang = float(xianYouShuLiang)

        ruKu = float(abs(box_value["{}_in".format(box_name)]))

        chuKu = float(abs(box_value["{}_out".format(box_name)]))

        chuShou = float(abs(box_value["{}_sale".format(box_name)]))

        value = float(xianYouShuLiang + ruKu - chuKu - chuShou)

    else:
        value = float(box_value["{}_god_hand".format(box_name)])

    return value

def parse_tabox(k_dict, google_auth, fernet_key, box_num, rule_df_dict, take_databases, book_dict, date_dict, wifi, backup_foldername):
    read_ = book_dict["read_"]
    dfb = date_dict["dfb"]
    yesterday = date_dict["yesterday"]

    tabox_inv_function = eval(str(k_dict["takeaway_box_inventory"]).strip().capitalize())

    if wifi:
        try:
            tabox_url = fernet_decrypt(str(k_dict["box_drink_in_out_url"]), fernet_key).strip()
            tabox_sheet = google_auth.open_by_url(tabox_url)

            tabox_sheet_name = str(k_dict["takeaway_box_sheetname"])
            tabox_sheet_name_index = tabox_sheet.worksheet(property="title",value=tabox_sheet_name).index
            tabox_df = tabox_sheet[tabox_sheet_name_index].get_as_df()

            #automatically update box names if different
            kdf_box_names = []
            for num in range(box_num):
                kdf_box_names += [str(k_dict["box{}".format(num)])]

            tabox_box_names = tabox_df.iloc[:, 1].to_list()

            if tabox_box_names != kdf_box_names:
                update_arr = np.array(kdf_box_names).reshape(len(kdf_box_names),1)

                update_matrix = []
                for l in update_arr:
                    update_matrix += [list(l)]

                dr = pygsheets.datarange.DataRange(start="B2", end="B41", worksheet=tabox_sheet[tabox_sheet_name_index])
                dr.update_values(values=update_matrix)
                tabox_df = tabox_sheet[tabox_sheet_name_index].get_as_df()
                print("BOX NAMES in {} had been updated automatically with accordance to the Rule.".format(tabox_sheet_name))
        except Exception as e:
            print()
            print("通过机器人从网络获取打包盒出入库数据失败，错误描述如下: ")
            print()
            print(e)
            print()

            local_db_filename = str(k_dict["local_database_filename"])
            tabox_sheet_name = str(k_dict["takeaway_box_sheetname"])

            tabox_df = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_db_filename), sheet_name=tabox_sheet_name)

            print("已采用本地备份的打包盒出入库数据。")
            print("注意确保本地备份的打包盒名字和数据库里的是一模一样的，当天所有出入库的数据是正确的。")
            print("如果需要更改本地备份数据，请先强制停止此程序，改好后重新运行。")
            print("如果无需更改任何数据, 按回车键继续运行程序")
            input("在这里输入>>>:")
    else:
        local_db_filename = str(k_dict["local_database_filename"])
        tabox_sheet_name = str(k_dict["takeaway_box_sheetname"])

        tabox_df = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_db_filename), sheet_name=tabox_sheet_name)

        print()
        print("无网络连接，已采用本地备份的打包盒出入库数据")
        print("注意确保本地备份的打包盒名字和数据库里是一模一样的，当天所有出入库的数据是正确的。")
        print("如果需要更改本地备份数据，请先强制停止此程序，改好后重新运行。")
        print("如果无需更改任何数据, 按回车键继续运行程序")
        input("在这里输入>>>:")

    if tabox_inv_function:
        tb_df = rule_df_dict["tb_df"]
        etb_df = rule_df_dict["etb_df"]

        tb_df["STRING LOCATOR"] = tb_df["STRING LOCATOR"].apply(lambda a: unicodedata.normalize("NFKD", a))
        tb_df["STRING LOCATOR"] = tb_df["STRING LOCATOR"].apply(lambda x: remove_spaces(x))
        tb_df["ACTIVE STATUS"] = tb_df["ACTIVE STATUS"].astype(int)

        model_listing = {}
        for number in range(box_num):
            model_name = k_dict["box{}".format(number)]
            criteria = (tb_df["MODEL TO USE"] == "box{}".format(number))&(tb_df["ACTIVE STATUS"] == 1)
            model_listing.update({ model_name : tb_df[criteria]["STRING LOCATOR"].values.tolist() })

        box_value = {}
        for index in range(len(tabox_df)):
            box_name = str(tabox_df.iloc[index, 1])
            in_stock = float(tabox_df.iloc[index, 2])*float(eval(str(tabox_df.iloc[index, 6])))
            out_stock = float(abs(tabox_df.iloc[index, 3]))*float(eval(str(tabox_df.iloc[index, 6])))
            god_hand  = float(tabox_df.iloc[index, 4])

            box_sales = box_use_sums(read_=read_,
                                       keyname=str(tabox_df.iloc[index, 1]),
                                       model_listing=model_listing,
                                       tb_df=tb_df,
                                       k_dict=k_dict)

            box_sales *= float(tabox_df.iloc[index, 5])

            box_sales *= float(eval(str(tabox_df.iloc[index, 6])))

            box_sales = float(normal_round(box_sales, 4))

            box_value.update({ "{}_in".format(box_name) : in_stock,
                               "{}_out".format(box_name) : out_stock,
                               "{}_god_hand".format(box_name) : god_hand,
                               "{}_sale".format(box_name) : box_sales,
                             })

        if etb_df.empty:
            pass

        else:
            etb_df["ACTIVE STATUS"] = etb_df["ACTIVE STATUS"].astype(int)

            for op in range(len(etb_df)):
                if int(etb_df["ACTIVE STATUS"][op]) != 1:
                    continue
                else:
                    operation = str(etb_df["OPERATION"][op]).strip()

                    if operation == "plus equal":
                        key0 = str(etb_df["BOX VALUE KEY TO OPERATE ON"][op])
                        key1 = str(etb_df["WITH KEY"][op])
                        box_value[key0] += box_value[key1]

                    elif operation == "minus equal":
                        key0 = str(etb_df["BOX VALUE KEY TO OPERATE ON"][op])
                        key1 = str(etb_df["WITH KEY"][op])
                        box_value[key0] -= box_value[key1]

                    elif operation == "divide equal":
                        key0 = str(etb_df["BOX VALUE KEY TO OPERATE ON"][op])
                        key1 = str(etb_df["WITH KEY"][op])
                        box_value[key0] /= box_value[key1]

                    elif operation == "times equal":
                        key0 = str(etb_df["BOX VALUE KEY TO OPERATE ON"][op])
                        key1 = str(etb_df["WITH KEY"][op])
                        box_value[key0] *= box_value[key1]

                    elif operation == "equal":
                        key0 = str(etb_df["BOX VALUE KEY TO OPERATE ON"][op])
                        key1 = str(etb_df["WITH KEY"][op])
                        box_value[key0] = box_value[key1]

                    elif operation == "zero":
                        key0 = str(etb_df["BOX VALUE KEY TO OPERATE ON"][op])
                        box_value[key0] = 0

                    elif operation == "change value":
                        key1 = float(etb_df["WITH KEY"][op])
                        box_value[key0] = key1

                    else:
                        pass

        takeaway_box_db_sheetname = k_dict["takeaway_box_database_sheetname"]
        tb_data = take_databases[takeaway_box_db_sheetname]

        tb_data["DATE"] = pd.to_datetime(tb_data["DATE"])
        dfb = pd.to_datetime(dfb)
        yesterday = pd.to_datetime(yesterday)

        tabox_ytd_record = tb_data[tb_data["DATE"] == yesterday]

        if tabox_ytd_record.empty:
            tabox_write_db = False
        else:
            tabox_write_db = True

        if tabox_write_db:
            for index in range(len(tabox_df)):
                box_name = str(tabox_df.iloc[index, 1])
                box_value_now = box_now(tb_data=tb_data,
                                        box_name=box_name,
                                        box_value=box_value,
                                        yesterday=yesterday)

                box_value.update({ "{}_now".format(box_name) : box_value_now })

            #update alert
            is_box_alert = eval(k_dict["box_stock_alert"].strip().capitalize())

            if is_box_alert:
                box_understock_alert = []
                box_understock_other = []
                box_stock_all = []

                box_overstock_alert = []

                for index in range(len(tabox_df)):
                    box_name = str(tabox_df.iloc[index, 1])

                    if box_name.startswith("box"):
                        continue
                    else:
                        box_value_now = float(box_value["{}_now".format(box_name)])
                        understock_alert_level = float(tabox_df.iloc[index, 7])
                        overstock_alert_level = float(tabox_df.iloc[index, 8])

                        box_stock_all += ["{}: 现有{}".format(tabox_df.iloc[index, 9], normal_round(box_value_now, 3))]

                        if box_value_now <= understock_alert_level:
                            box_understock_alert += ["{}: 现有{} (余量过少)".format(tabox_df.iloc[index, 9], normal_round(box_value_now, 3))]

                        else:
                            box_understock_other += ["{}: 现有{}".format(tabox_df.iloc[index, 9], normal_round(box_value_now, 3))]

                        if box_value_now >= overstock_alert_level:
                            box_overstock_alert += ["{}: 现有{} (余量过多)".format(tabox_df.iloc[index, 9], normal_round(box_value_now, 3))]

                        else:
                            continue
            else:
                box_understock_alert = []
                box_understock_other = []
                box_stock_all = []
                box_overstock_alert = []

        else:
            print("错误! 打包盒昨天的({})库存不存在数据库".format(yesterday.strftime("%Y-%m-%d")))
            box_understock_alert = []
            box_understock_other = []
            box_stock_all = []
            box_overstock_alert = []

    else:
        tabox_write_db = True

        box_value = {}
        for index in range(len(tabox_df)):
            box_name = str(tabox_df.iloc[index, 1])

            box_value.update({ "{}_in".format(box_name) : 0,
                               "{}_out".format(box_name) : 0,
                               "{}_god_hand".format(box_name) : 0,
                               "{}_sale".format(box_name) : 0,
                               "{}_now".format(box_name) : 0,
                             })

        box_understock_alert = []
        box_understock_other = []
        box_stock_all = []
        box_overstock_alert = []

    unparsed_alert = {
        "understock_alert": box_understock_alert,
        "understock_other" : box_understock_other,
        "stock_all" : box_stock_all,
        "overstock_alert": box_overstock_alert,
    }

    return tabox_write_db, box_value, unparsed_alert

def parse_alert(stock_alert_bool, date_dict, unparsed_alert, understock_alert_bool, overstock_alert_bool, alert_freq, alert_type, outlet):
    dfb = date_dict["dfb"]
    understock_alert = unparsed_alert["understock_alert"]
    understock_other = unparsed_alert["understock_other"]
    stock_all = unparsed_alert["stock_all"]
    overstock_alert = unparsed_alert["overstock_alert"]
    outlet = outlet.strip().capitalize()

    #box_stock_alert = eval(str(k_dict["box_stock_alert"]).strip().capitalize())
    #drink_stock_alert = eval(str(k_dict["drink_stock_alert"]).strip().capitalize())

    dfb = pd.to_datetime(dfb)

    if alert_type == "box":
        title_dict = {
            "stock_all" : "打包盒各型号余量: ",
            "understock_alert": "余量过少的打包盒型号: ",
            "understock_other" : "其他打包盒型号余量: ",
            "overstock_alert" : "余量过多的打包盒型号: ",
        }

    elif alert_type == "drink":
        title_dict = {
            "stock_all" : "各种饮品余量: ",
            "understock_alert": "余量过少的饮品: ",
            "understock_other": "其他饮品余量: ",
            "overstock_alert": "余量过多的饮品: ",
        }

    if stock_alert_bool:
        dayname = dfb.strftime("%A")
        dayname = dayname.strip().upper()

        #box_understock_alert = eval(str(k_dict["box_understocking_alert"]).strip().capitalize())
        #box_overstock_alert = eval(str(k_dict["box_overstocking_alert"]).strip().capitalize())
        #box_alert_freq = str(k_dict["box_stock_alert_frequency"]).strip().upper()

        #drink_understock_alert = eval(str(k_dict["drink_understocking_alert"]).strip().capitalize())
        #drink_overstock_alert = eval(str(k_dict["drink_overstocking_alert"]).strip().capitalize())
        #drink_alert_freq = str(k_dict["drink_stock_alert_frequency"]).strip().upper()

        if understock_alert_bool:
            if dayname == alert_freq:
                message_string = "{}({}) \n".format(outlet, dfb.strftime("%Y-%m-%d"))
                message_string += "{} \n".format(title_dict["stock_all"])
                for t in range(len(stock_all)):
                    message_string += "{} \n".format(stock_all[t])

            else:
                if len(understock_alert) > 0:
                    message_string = "{}({}) \n".format(outlet, dfb.strftime("%Y-%m-%d"))
                    message_string += "{} \n".format(title_dict["understock_alert"])
                    for t in range(len(understock_alert)):
                        message_string += "{} \n".format(understock_alert[t])

                    message_string += "------------------------ \n"
                    message_string += "{} \n".format(title_dict["understock_other"])

                    for s in range(len(understock_other)):
                        message_string += "{} \n".format(understock_other[s])

                else:
                    message_string = ""
        else:
            message_string = ""


        if overstock_alert_bool:
            try:
                message_string+"" == message_string
            except:
                message_string = "{}({}) \n".format(outlet, dfb.strftime("%Y-%m-%d"))

            if len(overstock_alert) > 0:
                message_string += "------------------------ \n"
                message_string += "{} \n".format(title_dict["overstock_alert"])
                for t in range(len(overstock_alert)):
                    message_string += "{} \n".format(overstock_alert[t])

            else:
                pass
        else:
            pass
    else:
        message_string = ""


    return message_string

def sending_email(is_pr, mail_server, mail_sender, mail_sender_password, mail_receivers, mail_subject, message_string, wifi):
    if wifi:
        if len(message_string) == 0:
            pass
        else:
            if is_pr:
                msg_str = ""
                for text in message_string:
                    msg_str += "{} \n".format(text)
            else:
                msg_str = message_string

            msg = MIMEText(msg_str, "plain", "utf-8")
            msg["From"] = mail_sender

            msg["Subject"] = Header(mail_subject, "utf-8")

            try:
                smtp_obj = smtplib.SMTP()
                smtp_obj.connect(mail_server, 25)
                smtp_obj.login(mail_sender, mail_sender_password)

                print("正在发送电邮给{}".format(mail_receivers))
                msg["To"] = mail_receivers
                smtp_obj.sendmail(mail_sender, mail_receivers, msg.as_string())

                print("{} 电邮发送成功".format(mail_receivers))

            except smtplib.SMTPException:
                print("{} 电邮发送失败".format(mail_receivers))
    else:
        print("无网络连接，无法发送电邮")

def look_up_rule(dataFrame, a_string, take):
    df = dataFrame.copy()
    df['0'] = df['0'].astype(str)

    if int(take) == 1:
        dfSlice = df[df['0'].str.contains(a_string)].dropna(axis=1)

    elif int(take) == 2:
        dfSlice = df[df['0'].str.contains(a_string,regex=False)].dropna(axis=1)

    elif int(take) == 3:
        dfSlice = df[df['0'] == a_string].dropna(axis=1)

    elif int(take) == 4:
        dfSlice = df[df['0'].str.contains(a_string, case=False, regex=False)].dropna(axis=1)

    else:
        dfSlice = pd.dataFrame(columns=df.columns)

    return dfSlice

def sum_rule(dfSlice, take, is_money):
    if dfSlice.empty:
        value = 0

    else:

        dfsc = np.arange(0, len(dfSlice.columns)).astype(str)
        dfSlice.columns = dfsc

        if int(take) == 1:
            value = float(dfSlice['1'].values[0])

        elif int(take) == 2:
            dfSlice['1'] = dfSlice['1'].astype(float)
            value = float(dfSlice['1'].sum())

        elif int(take) == 3:
            a_list = []
            eval_a_list = []

            for x in dfSlice['0'].values:
                res = re.findall(r'\(.*?\)', x)
                a_list += res

            for y in range(len(a_list)):
                eval_a_list += [eval(a_list[y])]

            value = sum(eval_a_list)

        elif int(take) == 4:
            dfSlice['1'] = dfSlice['1'].astype(int)
            value = int(dfSlice['1'].sum())

        else:
            value = -404

    if int(is_money) == 1:
        value = single_zero(format(value, '.2f'))
    else:
        pass

    return value

def promo_add_rule(take_databases, k_dict, dfb, yesterday, promo_suffix, promo_count):
    promotion_db_sheetname = k_dict["promotion_database_sheetname"]
    promo_db = take_databases[promotion_db_sheetname]

    promo_db["DATE"] = pd.to_datetime(promo_db["DATE"])
    dfb = pd.to_datetime(dfb)
    yesterday = pd.to_datetime(yesterday)

    promo_ytd_record = promo_db[promo_db["DATE"] == yesterday]

    promo_suffix = promo_suffix.replace("_count", "")

    if promo_ytd_record.empty:
        promo_write_db = False
    else:
        promo_write_db = True

    if promo_write_db:
        promo_reset = str(k_dict["{}_reset_count".format(promo_suffix)])
        dfb = pd.to_datetime(dfb)
        dayname = dfb.strftime("%A")
        dayname = dayname.strip().upper()

        if promo_reset == "FALSE":
            promo_count_ytd = int(promo_db[promo_db["DATE"] == yesterday]["{}_tdy".format(promo_suffix)].values[0])
            promo_count_tdy = promo_count_ytd + int(promo_count)

        elif promo_reset in ["TODAY", "EVERYDAY"]:
            promo_count_ytd = 0
            promo_count_tdy = promo_count_ytd + int(promo_count)

        elif promo_reset == "START OF MONTH":
            if int(dfb.day) == 1:
                promo_count_ytd = 0
                promo_count_tdy = promo_count_ytd + int(promo_count)

            else:
                promo_count_ytd = int(promo_db[promo_db["DATE"] == yesterday]["{}_tdy".format(promo_suffix)].values[0])
                promo_count_tdy = promo_count_ytd + int(promo_count)
        else:
            if dayname == promo_reset:
                promo_count_ytd = 0
                promo_count_tdy = promo_count_ytd + int(promo_count)

            else:
                promo_count_ytd = int(promo_db[promo_db["DATE"] == yesterday]["{}_tdy".format(promo_suffix)].values[0])
                promo_count_tdy = promo_count_ytd + int(promo_count)

        return_list = [promo_count_ytd, promo_count_tdy]
    else:
        return_list = [-1000000, -1000000]

    return return_list

def parse_value_dict(promo_num, lun_sales, lun_gc, tb_sales, tb_gc, lun_fwc, lun_kwc, tb_fwc, tb_kwc, night_fwc, night_kwc, rule_df_dict, take_databases, date_dict, book_dict, k_dict):

    read = book_dict["read"]
    is_local_book = book_dict["is_local_book"]
    outlet_loc = book_dict["outlet_loc"]

    dfb = date_dict["dfb"]
    yesterday = date_dict["yesterday"]
    date_chinese_string = date_dict["date_chinese_string"]

    total_sales = read[read['0'] == 'Total Sales'].dropna(axis=1)
    tsc = np.arange(0, len(total_sales.columns)).astype(str)
    total_sales.columns = tsc
    total_sales = float(total_sales["1"].values[0])

    no_of_cover = read[read['0'] == 'No. Of Cover'].dropna(axis=1)
    nocc = np.arange(0, len(no_of_cover.columns)).astype(str)
    no_of_cover.columns = nocc
    no_of_cover = int(no_of_cover["1"].values[0])

    if is_local_book:
        lun_start_index = int(read[read['0'].str.contains('LUNCH')].index[0])
        lun_end_index = int(read[read['0'].str.contains('TEA TIME')].index[0])
        lunch_df = read.copy()
        lunch_df = lunch_df.iloc[(lun_start_index+1):(lun_end_index), :]
        lunch_df.dropna(axis=1, inplace=True)

        lun_sales = float(lunch_df[lunch_df['0'].str.contains('Total Amount')]['1'].values[0])
        lun_avg = float(lunch_df[lunch_df['0'].str.contains('Average/Pax')]['1'].values[0])
        lun_gc = int(lunch_df[lunch_df['0'].str.contains('Total Pax')]['1'].values[0])

    else:
        lun_sales = float(lun_sales)

        if lun_gc == 0:
            lun_avg = 0
        else:
            lun_avg = lun_sales/int(lun_gc)

        lun_avg = normal_round(lun_avg, 2)

    if is_local_book:
        tb_start_index = lun_end_index
        tb_end_index = int(read[read['0'].str.contains('DINNER')].index[0])
        tb_df = read.copy()
        tb_df = tb_df.iloc[(tb_start_index)+1:(tb_end_index), :]
        tb_df.dropna(axis=1, inplace=True)

        tb_sales = float(tb_df[tb_df['0'].str.contains('Total Amount')]['1'].values[0])
        tb_avg = float(tb_df[tb_df['0'].str.contains('Average/Pax')]['1'].values[0])
        tb_gc = int(tb_df[tb_df['0'].str.contains('Total Pax')]['1'].values[0])

    else:
        tb_sales = float(tb_sales)

        if tb_gc == 0:
            tb_avg = 0
        else:
            tb_avg = tb_sales/int(tb_gc)

        tb_avg = normal_round(tb_avg, 2)

    if is_local_book:
        night_start_index = tb_end_index
        night_end_index = night_start_index+6
        night_df = read.copy()
        night_df = night_df.iloc[(night_start_index+1):(night_end_index), :]
        night_df.dropna(axis=1, inplace=True)

        night_sales = float(night_df[night_df['0'].str.contains('Total Amount')]['1'].values[0])
        night_avg = float(night_df[night_df['0'].str.contains('Average/Pax')]['1'].values[0])
        night_gc = int(night_df[night_df['0'].str.contains('Total Pax')]['1'].values[0])

    else:
        night_sales = total_sales - lun_sales - tb_sales
        night_gc = no_of_cover - int(lun_gc) - int(tb_gc)

        if night_gc == 0:
            night_avg = 0
        else:
            night_avg = night_sales/int(night_gc)

        night_sales = normal_round(night_sales, 2)
        night_avg = normal_round(night_avg, 2)

    lun_sales = single_zero(format(lun_sales, '.2f'))
    lun_avg = single_zero(format(lun_avg, '.2f'))
    tb_sales = single_zero(format(tb_sales, '.2f'))
    tb_avg = single_zero(format(tb_avg, '.2f'))
    night_sales = single_zero(format(night_sales, '.2f'))
    night_avg = single_zero(format(night_avg, '.2f'))

    value_dict = {
        'date_chinese_string' : date_chinese_string,
        'outlet_loc' : outlet_loc,
        'lun_sales' : lun_sales,
        'lun_gc' : lun_gc,
        'lun_avg' : lun_avg,
        'lun_fwc' : lun_fwc,
        'lun_kwc' : lun_kwc,
        'tb_sales' : tb_sales,
        'tb_gc' : tb_gc,
        'tb_avg' : tb_avg,
        'tb_fwc': tb_fwc,
        'tb_kwc' : tb_kwc,
        'night_sales' : night_sales,
        'night_gc' : night_gc,
        'night_avg' : night_avg,
        'night_fwc' : night_fwc,
        'night_kwc' : night_kwc,
        'no_of_cover' : no_of_cover,
    }

    vd_df = rule_df_dict["vd_df"]

    net_sales_deduct_total = 0
    for index in range(len(vd_df)):
        if int(vd_df.iloc[index, 8]) == 1:
            dataframe = book_dict["{}".format(vd_df.iloc[index, 2])]
            a_string = str(vd_df.iloc[index, 1])
            look_up_rule_take = int(vd_df.iloc[index, 3])

            df_slice = look_up_rule(dataFrame=dataframe, a_string=a_string, take=look_up_rule_take)

            sum_rule_take = int(vd_df.iloc[index, 4])
            is_money = int(vd_df.iloc[index, 5])

            key_value = sum_rule(dfSlice=df_slice, take=sum_rule_take, is_money=is_money)
            key = str(vd_df.iloc[index, 0])

            value_dict.update({ key : key_value })

            if key.startswith("promo") and key.endswith("_count"):
                promo_key = "{}_ytd_tdy".format(key.replace("_count", ""))

                promo_add_rule_list = promo_add_rule(take_databases=take_databases,
                                                     k_dict=k_dict,
                                                     dfb=dfb,
                                                     yesterday=yesterday,
                                                     promo_suffix=key,
                                                     promo_count=value_dict[key])

                value_dict.update({ promo_key : promo_add_rule_list })

            if int(vd_df.iloc[index, 6]) == 1:
                net_sales_deduct_total += float(key_value)

        else:
            continue

    value_dict.update({ "net_sales_after_deduction" : float(float(value_dict["net_sales"]) - net_sales_deduct_total) })
    value_dict["net_sales_after_deduction"] = single_zero(format(normal_round(value_dict["net_sales_after_deduction"], 2), ".2f"))

    financial_database_sheetname = k_dict["financial_database_sheetname"]
    financial_data = take_databases[financial_database_sheetname]

    financial_data["DATE"] = pd.to_datetime(financial_data["DATE"])
    dfb = pd.to_datetime(dfb)
    yesterday = pd.to_datetime(yesterday)

    financial_ytd_record = financial_data[financial_data["DATE"] == yesterday]

    if financial_ytd_record.empty:
        write_finance_db = False
    else:
        write_finance_db = True

    if write_finance_db:
        if int(dfb.day) == 1:
            cmns_ytd = 0.0
            cmns_tdy = cmns_ytd + float(value_dict["net_sales_after_deduction"])
            cmns_tdy = float(format(normal_round(cmns_tdy, 2), ".2f"))
        else:
            cmns_ytd = float(financial_ytd_record["CUM NET SALES TODAY"].values[0])
            cmns_tdy = cmns_ytd + float(value_dict["net_sales_after_deduction"])
            cmns_tdy = float(format(normal_round(cmns_tdy, 2), ".2f"))

        avg_daily_sales = cmns_tdy/int(dfb.day)
        ads_rd1 = str(avg_daily_sales)[str(avg_daily_sales).find('.'):][:3]
        ads_rd2 = str(avg_daily_sales)[:str(avg_daily_sales).find('.')]
        ads = float(ads_rd2+ads_rd1)

        cmns_tdy = format(cmns_tdy, ".2f")

        value_dict.update({
            "cmns_ytd" : cmns_ytd,
            "cmns_tdy" : cmns_tdy,
            "ads" : ads,
        })

    else:
        print("错误! {}的昨天({})的营业额信息不存在".format(dfb.strftime("%Y-%m-%d"), yesterday.strftime("%Y-%m-%d")))
        value_dict.update({
            "cmns_ytd" : -1000000.00,
            "cmns_tdy" : -1000000.00,
            "ads" : -1000000.00,
        })

    for num in range(promo_num):
        value_dict.update({ "promo{}_ytd".format(num+1) : value_dict["promo{}_ytd_tdy".format(num+1)][0],
                            "promo{}_tdy".format(num+1) : value_dict["promo{}_ytd_tdy".format(num+1)][1],
                          })

    promotion_database_sheetname = k_dict["promotion_database_sheetname"]
    promo_data = take_databases[promotion_database_sheetname]

    promo_data["DATE"] = pd.to_datetime(promo_data["DATE"])
    dfb = pd.to_datetime(dfb)
    yesterday = pd.to_datetime(yesterday)

    promo_ytd_record = promo_data[promo_data["DATE"] == yesterday]

    if promo_ytd_record.empty:
        write_promo_db = False
    else:
        write_promo_db = True

    return value_dict, write_finance_db, write_promo_db

def feedback_parser(k_dict):
    food_filename = str(k_dict["food_feedback_filename"])
    food_default = str(k_dict["food_feedback_default_text"])

    if os.path.exists(food_filename):
        with open(food_filename, "r") as file_handler:
            food_feedback = file_handler.read()

        food_feedback = food_feedback.split("\n")

    else:
        with open(food_filename, "w") as file_handler:
            file_handler.write(food_default)

        with open(food_filename, "r") as file_handler:
            food_feedback = file_handler.read()

        food_feedback = food_feedback.split("\n")
        print("'{}'文件不存在, 菜肴反馈已采用默认文本。".format(food_filename))

    service_filename = str(k_dict["service_feedback_filename"])
    service_default = str(k_dict["service_feedback_default_text"])

    if os.path.exists(service_filename):
        with open(service_filename, "r") as file_handler:
            service_feedback = file_handler.read()

        service_feedback = service_feedback.split("\n")

    else:
        with open(service_filename, "w") as file_handler:
            file_handler.write(service_default)

        with open(service_filename, "r") as file_handler:
            service_feedback = file_handler.read()

        service_feedback = service_feedback.split("\n")
        print("'{}'文件不存在, 服务反馈已采用默认文本。".format(service_filename))

    return food_feedback, service_feedback

def parse_print_rule(value_dict, rule_df_dict, write_finance_db, write_promo_db, write_drink_db, tabox_write_db, k_dict):

    food_feedback, service_feedback = feedback_parser(k_dict=k_dict)

    pr_df = rule_df_dict["pr_df"]

    for col in pr_df.columns:
        pr_df[col] = pr_df[col].astype(str)

    print_result = []
    for index in range(len(pr_df)):
        if int(pr_df["ACTIVE STATUS"][index]) == 1:
            lis = pr_df["KEY IN VALUE DICT"][index]
            listt = lis.split(",")
            f_eval = []

            for f in listt:
                if f == "nan":
                    continue

                elif f == "":
                    continue

                else:
                    f_eval += [value_dict[str(f)]]

            if pr_df["PRINT STATEMENT"][index] == "nan":
                print_result += [" "]

            elif pr_df["PRINT STATEMENT"][index] == "":
                print_result += [" "]

            else:
                print_result += [pr_df["PRINT STATEMENT"][index].format(*f_eval)]
        else:
            continue

    print_result += ["服务反馈: "]

    for line in service_feedback:
        print_result += [line]

    print_result += [" ", "菜肴反馈: "]

    for line in food_feedback:
        print_result += [line]

    print_error = False
    for booleans in [write_finance_db,write_promo_db,write_drink_db,tabox_write_db]:
        if not booleans:
            print_error = True
        else:
            continue

    if print_error:
        print("注意！由于有些数据的缺失，部分显示的数字可能不会准确！")

    return print_result

def feedback_reset(k_dict, reset_remarks):
    food_default = str(k_dict["food_feedback_default_text"])
    service_default = str(k_dict["service_feedback_default_text"])

    food_filename = str(k_dict["food_feedback_filename"])
    service_filename = str(k_dict["service_feedback_filename"])

    if reset_remarks:
        with open(service_filename, "w") as file_handler:
            file_handler.write(service_default)

        print("服务反馈已自动取代为默认文本。")
        print()

        with open(food_filename, "w") as file_handler:
            file_handler.write(food_default)

        print("菜肴反馈已自动取代为默认文本。")
        print()

    else:
        print()
        print("是否重置服务反馈为默认文本?")
        print()
        action_req = option_num(["是", "否"])
        time.sleep(0.25)
        user_input = option_limit(action_req, input("在这里输入>>>: "))

        if user_input == 0:
            with open(service_filename, "w") as file_handler:
                file_handler.write(service_default)

            print("服务反馈已取代为默认文本。")

        else:
            print("好的，服务反馈没有被重置。")

        print("是否重置菜肴反馈为默认文本?")
        print()
        action_req = option_num(["是", "否"])
        time.sleep(0.25)
        user_input = option_limit(action_req, input("在这里输入>>>: "))

        if user_input == 0:
            with open(food_filename, "w") as file_handler:
                file_handler.write(food_default)

            print("菜肴反馈已取代为默认文本。")

        else:
            print("好的，菜肴反馈没有被重置。")

def drink_remarks(drink_out, drink_inv_function):
    if drink_inv_function:
        if drink_out == 0:
            return None
        else:
            return "自用{}".format(int(drink_out))
    else:
        return None

def parse_drink_stock(k_dict, fernet_key, google_auth, drink_num, value_dict, take_databases, date_dict, wifi, backup_foldername):
    dfb = date_dict["dfb"]
    yesterday = date_dict["yesterday"]

    drink_inv_function = eval(str(k_dict["drink_inventory"]).strip().capitalize())

    if wifi:
        try:
            drink_url = fernet_decrypt(str(k_dict["box_drink_in_out_url"]), fernet_key).strip()
            drink_sheet = google_auth.open_by_url(drink_url)

            drink_stock_sheetname = str(k_dict["drink_stock_sheetname"])
            drink_stock_sheetname_index = drink_sheet.worksheet(property="title",value=drink_stock_sheetname).index
            drink_df = drink_sheet[drink_stock_sheetname_index].get_as_df()

            #automatically update drink names if different
            kdf_drink_names = []
            for num in range(drink_num):
                kdf_drink_names += [str(k_dict["drink{}".format(num)])]

            drinkdf_drink_name = drink_df.iloc[:, 1].to_list()

            if drinkdf_drink_name != kdf_drink_names:
                update_arr = np.array(kdf_drink_names).reshape(len(kdf_drink_names), 1)

                update_matrix = []
                for l in update_arr:
                    update_matrix += [list(l)]

                dr = pygsheets.datarange.DataRange(start="B2", end="B13", worksheet=drink_sheet[drink_stock_sheetname_index])
                dr.update_values(values=update_matrix)
                drink_df = drink_sheet[drink_stock_sheetname_index].get_as_df()
                print("DRINK NAMES in {} had been updated automatically with accordance to the Rule".format(drink_stock_sheetname))
        except Exception as e:
            print()
            print("通过机器人从网络获取酒水出入库数据失败，错误描述如下: ")
            print(e)
            print()
            local_db_filename = str(k_dict["local_database_filename"])
            drink_stock_sheetname = str(k_dict["drink_stock_sheetname"])

            drink_df = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_db_filename), sheet_name=drink_stock_sheetname)

            print("已采用本地备份的酒水出入库数据。")
            print("注意确保本地备份的酒水名字和数据库里的是一模一样的，当天所有出入库的数据是正确的。")
            print("如果需要更改本地备份数据，请先强制停止此程序，改好后重新运行。")
            print("如果无需更改任何数据, 按回车键继续运行程序")
            input("在这里输入>>>:")
    else:
        local_db_filename = str(k_dict["local_database_filename"])
        drink_stock_sheetname = str(k_dict["drink_stock_sheetname"])

        drink_df = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_db_filename), sheet_name=drink_stock_sheetname)

        print("无网络连接，已采用本地备份的酒水出入库数据")
        print("注意确保本地备份的酒水名字和数据库里的是一模一样的，当天所有出入库的数据是正确的。")
        print("如果需要更改本地备份数据，请先强制停止此程序，改好后重新运行。")
        print("如果无需更改任何数据, 按回车键继续运行程序")
        input("在这里输入>>>:")

    if drink_inv_function:

        drink_db_sheetname = k_dict["drink_database_sheetname"]
        drink_data = take_databases[drink_db_sheetname]

        drink_data["DATE"] = pd.to_datetime(drink_data["DATE"])
        dfb = pd.to_datetime(dfb)
        yesterday = pd.to_datetime(yesterday)

        drink_data_ytd_record = drink_data[drink_data["DATE"] == yesterday]

        if drink_data_ytd_record.empty:
            write_drink_db = False
        else:
            write_drink_db = True

        if write_drink_db:
            drink_dict = {}
            for index in range(len(drink_df)):
                drink_index = str(drink_df.iloc[index, 0])
                drink_name = str(drink_df.iloc[index, 1])
                drink_sale = int(value_dict["{}_sale".format(drink_index)])
                drink_in = int(drink_df.iloc[index, 2])
                drink_out = int(drink_df.iloc[index, 3])
                drink_previous = int(drink_data[drink_data["DATE"] == yesterday]["{}实结存".format(drink_name)].values[0])
                drink_memo = drink_remarks(drink_out = drink_out, drink_inv_function=drink_inv_function)

                drink_dict.update({
                    "{}上日存货".format(drink_name) : drink_previous ,
                    "{}进".format(drink_name) : drink_in ,
                    "{}出".format(drink_name) : abs(drink_out) + abs(drink_sale) ,
                    "{}实结存".format(drink_name) : (drink_previous + drink_in - drink_out - drink_sale) ,
                    "{}备注".format(drink_name) : drink_memo ,
                })

            #update alert list
            is_drink_alert = eval(k_dict["drink_stock_alert"].strip().capitalize())

            if is_drink_alert:
                drink_understock_alert = []
                drink_understock_other = []
                drink_stock_all = []

                drink_overstock_alert = []

                for index in range(len(drink_df)):
                    drink_name = str(drink_df.iloc[index, 1])

                    if drink_name.startswith("饮料"):
                        continue

                    else:
                        drink_value_now = drink_dict["{}实结存".format(drink_name)]
                        understock_alert_level = float(drink_df.iloc[index, 4])
                        overstock_alert_level = float(drink_df.iloc[index, 5])

                        drink_stock_all += ["{}: 现有{}".format(drink_df.iloc[index, 6], drink_value_now)]

                        if drink_value_now <= understock_alert_level:
                            drink_understock_alert += ["{}: 现有{}(余量过少)".format(drink_df.iloc[index,6], drink_value_now)]

                        else:
                            drink_understock_other += ["{}: 现有{}".format(drink_df.iloc[index,6], drink_value_now)]

                        if drink_value_now >= overstock_alert_level:
                            drink_overstock_alert += ["{}: 现有{}(余量过多)".format(drink_df.iloc[index, 6], drink_value_now)]
            else:
                drink_understock_alert = []
                drink_understock_other = []
                drink_stock_all = []

                drink_overstock_alert = []
        else:
            print("错误！酒水昨天({})的库存不存在! ".format(yesterday.strftime("%Y-%m-%d")))
            drink_dict = {}
            drink_understock_alert = []
            drink_understock_other = []
            drink_stock_all = []

            drink_overstock_alert = []
    else:
        write_drink_db = True
        drink_dict = {}
        for index in range(len(drink_df)):
            drink_name = drink_df.iloc[index, 1]
            drink_memo = drink_remarks(drink_out = 0, drink_inv_function=drink_inv_function)

            drink_dict.update({
                    "{}上日存货".format(drink_name): 0,
                    "{}进".format(drink_name): 0,
                    "{}出".format(drink_name): 0,
                    "{}实结存".format(drink_name): 0,
                    "{}备注".format(drink_name): drink_memo,
            })

        drink_understock_alert = []
        drink_understock_other = []
        drink_stock_all = []

        drink_overstock_alert = []

    unparsed_alert = {
        "understock_alert" : drink_understock_alert,
        "understock_other" : drink_understock_other,
        "stock_all" : drink_stock_all,
        "overstock_alert" : drink_overstock_alert,
    }

    return write_drink_db, drink_dict, unparsed_alert

def db_write(write_drink_db, tabox_write_db, write_finance_db, write_promo_db):
    db_writables = {
        "write_drink_db": write_drink_db,
        "tabox_write_db": tabox_write_db,
        "write_finance_db" : write_finance_db,
        "write_promo_db" : write_promo_db,
    }

    return db_writables

def pending_upload_db(take_databases, k_dict, box_value, drink_dict, value_dict, db_writables, date_dict, drink_num, promo_num, get_columns):
    #print("正在处理待上传的数据...")
    #print()

    drink_db_sheetname = k_dict["drink_database_sheetname"]
    takeaway_box_db_sheetname = k_dict["takeaway_box_database_sheetname"]
    financial_db_sheetname = k_dict["financial_database_sheetname"]
    promotion_db_sheetname = k_dict["promotion_database_sheetname"]

    sheet_names = [drink_db_sheetname, takeaway_box_db_sheetname,
                   financial_db_sheetname, promotion_db_sheetname]


    dfb = date_dict["dfb"]
    yesterday = date_dict["dfb"]
    date_chinese_string = date_dict["date_chinese_string"]

    dfb = pd.to_datetime(dfb)
    yesterday = pd.to_datetime(yesterday)

    write_drink_db = db_writables["write_drink_db"]
    tabox_write_db = db_writables["tabox_write_db"]
    write_finance_db = db_writables["write_finance_db"]
    write_promo_db = db_writables["write_promo_db"]

    drink_col = get_columns["drink_col"]
    tabox_col = get_columns["tabox_col"]
    financial_col = get_columns["financial_col"]
    promo_col = get_columns["promo_col"]

    #update drink database
    if write_drink_db:
        curr_db = take_databases[sheet_names[0]]
        curr_db["DATE"] = pd.to_datetime(curr_db["DATE"])

        tdy_record = curr_db[curr_db["DATE"] == dfb]

        if tdy_record.empty:
            riQi = date_chinese_string
            date = dfb
            time_log = pd.to_datetime(dt.datetime.now())

            u_list = [riQi,]
            for n in range(drink_num):
                drink_name = k_dict["drink{}".format(n)]
                u_list += [
                    drink_dict["{}上日存货".format(drink_name)],
                    drink_dict["{}进".format(drink_name)],
                    drink_dict["{}出".format(drink_name)],
                    drink_dict["{}实结存".format(drink_name)],
                    drink_dict["{}备注".format(drink_name)],
                ]

            u_list += [date, time_log]

            u_dict = {}
            for index in range(len(drink_col)):
                u_dict.update({ drink_col[index] : [u_list[index]] })

            u_dict = pd.DataFrame(u_dict)
            u_dict["DATE"] = pd.to_datetime(u_dict["DATE"])

            drink_concat = pd.concat([curr_db, u_dict])
            drink_concat["DATE"] = pd.to_datetime(drink_concat["DATE"])
            drink_concat.sort_values(by=["DATE"], ascending=False, ignore_index=True, inplace=True)
            take_databases[sheet_names[0]] = drink_concat
            #print("酒水数据处理成功! {}的酒水数据已加入待上传列队".format(dfb.strftime("%Y-%m-%d")))
            send_drink_msg = True

        else:
            print("{}的酒水数据已存在数据库，不能覆盖原数据".format(dfb.strftime("%Y-%m-%d")))
            send_drink_msg = False
    else:
        print("酒水数据处理失败!")
        send_drink_msg = False

    #update takeaway boxes database
    if tabox_write_db:
        curr_db = take_databases[sheet_names[1]]
        curr_db["DATE"] = pd.to_datetime(curr_db["DATE"])

        tdy_record = curr_db[curr_db["DATE"] == dfb]

        if tdy_record.empty:
            time_log = pd.to_datetime(dt.datetime.now())
            date = dfb
            u_list = [time_log, date]

            for index in range(box_num):
                box_name = k_dict["box{}".format(index)]
                box_outt = float(box_value["{}_out".format(box_name)]) + float(box_value["{}_sale".format(box_name)])

                u_list += [
                           float(box_value["{}_in".format(box_name)]),
                           box_outt,
                           float(box_value["{}_now".format(box_name)]),
                          ]

            u_dict = {}
            for index in range(len(tabox_col)):
                u_dict.update({ tabox_col[index] : [u_list[index]] })

            u_dict = pd.DataFrame(u_dict)
            u_dict["DATE"] = pd.to_datetime(u_dict["DATE"])

            tabox_concat = pd.concat([curr_db, u_dict])
            tabox_concat["DATE"] = pd.to_datetime(tabox_concat["DATE"])
            tabox_concat.sort_values(by=["DATE"], ascending=False, ignore_index=True, inplace=True)
            take_databases[sheet_names[1]] = tabox_concat
            #print("打包盒数据处理成功! {}的打包盒数据已加入待上传列队".format(dfb.strftime("%Y-%m-%d")))
            send_tabox_msg = True
        else:
            print("{}的打包盒数据已存在数据库，不能覆盖原数据".format(dfb.strftime("%Y-%m-%d")))
            send_tabox_msg = False
    else:
        print("打包盒数据处理失败! ")
        send_tabox_msg = False

    if write_finance_db:
        curr_db = take_databases[sheet_names[2]]
        curr_db["DATE"] = pd.to_datetime(curr_db["DATE"])

        tdy_record = curr_db[curr_db["DATE"] == dfb]

        if tdy_record.empty:
            time_log = pd.to_datetime(dt.datetime.now())
            date = dfb
            day = dfb.strftime("%A")
            net_sales = value_dict["net_sales_after_deduction"]
            cum_net_sales_ytd = value_dict["cmns_ytd"]
            cum_net_sales_tdy = float(value_dict["cmns_tdy"])
            avg_sales_per_day = value_dict["ads"]
            cash_amt = value_dict["CASH"]
            dine_in_total = value_dict["SALES_DI"]
            takeaway_total = value_dict["SALES_TA"]

            u_list = [time_log, date, day,
                      net_sales, cum_net_sales_ytd, cum_net_sales_tdy,
                      avg_sales_per_day, cash_amt,
                      dine_in_total, takeaway_total]

            u_dict = {}
            for index in range(len(financial_col)):
                u_dict.update({ financial_col[index] : [u_list[index]] })

            u_dict = pd.DataFrame(u_dict)
            u_dict["DATE"] = pd.to_datetime(u_dict["DATE"])

            finance_concat = pd.concat([curr_db, u_dict])
            finance_concat["DATE"] = pd.to_datetime(finance_concat["DATE"])
            finance_concat.sort_values(by=["DATE"], ascending=False, ignore_index=True, inplace=True)
            take_databases[sheet_names[2]] = finance_concat
            #print("财务数据处理成功! {}的财务数据已加入待上传列队".format(dfb.strftime("%Y-%m-%d")))
            send_print_result = True
        else:
            print("{}的财务数据已存在数据库，不能覆盖原数据".format(dfb.strftime("%Y-%m-%d")))
            send_print_result = False
    else:
        print("财务数据处理失败! ")
        send_print_result = False

    if write_promo_db:
        curr_db = take_databases[sheet_names[3]]
        curr_db["DATE"] = pd.to_datetime(curr_db["DATE"])

        tdy_record = curr_db[curr_db["DATE"] == dfb]

        if tdy_record.empty:
            time_log = pd.to_datetime(dt.datetime.now())
            date = dfb

            u_list = [time_log, date]
            for i in range(promo_num):
                u_list += [value_dict["promo{}_ytd".format(i+1)],
                           value_dict["promo{}_tdy".format(i+1)],
                          ]
            u_dict = {}
            for index in range(len(promo_col)):
                u_dict.update({ promo_col[index] : [u_list[index]] })

            u_dict = pd.DataFrame(u_dict)
            u_dict["DATE"] = pd.to_datetime(u_dict["DATE"])

            promo_concat = pd.concat([curr_db, u_dict])
            promo_concat["DATE"] = pd.to_datetime(promo_concat["DATE"])
            promo_concat.sort_values(by=["DATE"], ascending=False, ignore_index=True, inplace=True)
            take_databases[sheet_names[3]] = promo_concat
            #print("促销数据处理成功! {}的促销数据已加入待上传列队".format(dfb.strftime("%Y-%m-%d")))
        else:
            print("{}的促销数据已存在数据库，不能覆盖原数据".format(dfb.strftime("%Y-%m-%d")))
            send_print_result = False
    else:
        print("促销数据处理失败! ")
        send_print_result = False

    send_dict = {
        "send_drink_msg" : send_drink_msg,
        "send_tabox_msg" : send_tabox_msg,
        "send_print_result" : send_print_result,
    }

    return take_databases, send_dict

def upload_db(database_url, take_databases, k_dict, fernet_key, google_auth, box_num, drink_num, wifi, backup_foldername):
    #print("正在上传数据...")

    auto_hour = 20
    auto_min = 55

    NOW = dt.datetime.now()
    AUTO_TIME = dt.datetime(NOW.year, NOW.month, NOW.day, auto_hour, auto_min)

    local_db_filename = str(k_dict["local_database_filename"])
    online_database_url = database_url
    online_database_url = fernet_decrypt(online_database_url, fernet_key)

    drink_db_sheetname = k_dict["drink_database_sheetname"]
    takeaway_box_db_sheetname = k_dict["takeaway_box_database_sheetname"]
    financial_db_sheetname = k_dict["financial_database_sheetname"]
    promotion_db_sheetname = k_dict["promotion_database_sheetname"]
    shift_db_sheetname = k_dict["shift_database_sheetname"]

    sheet_names = [drink_db_sheetname, takeaway_box_db_sheetname,
                   financial_db_sheetname, promotion_db_sheetname]

    tabox_url = str(k_dict["box_drink_in_out_url"])
    tabox_url = fernet_decrypt(tabox_url, fernet_key).strip()

    takeaway_box_sheetname = str(k_dict["takeaway_box_sheetname"])
    drink_stock_sheetname = str(k_dict["drink_stock_sheetname"])
    receivers_sheetname = str(k_dict["receivers_sheetname"])

    if wifi:
        try:
            tabox_drink_sheet = google_auth.open_by_url(tabox_url)

            tabox_sheet_index = tabox_drink_sheet.worksheet(property="title", value=takeaway_box_sheetname).index
            drink_sheet_index = tabox_drink_sheet.worksheet(property="title", value=drink_stock_sheetname).index

            tabox_df = tabox_drink_sheet[tabox_sheet_index].get_as_df()
            drink_df = tabox_drink_sheet[drink_sheet_index].get_as_df()

            dr1 = pygsheets.datarange.DataRange(start="C2", end="C41", worksheet=tabox_drink_sheet[tabox_sheet_index])
            dr2 = pygsheets.datarange.DataRange(start="D2", end="D41", worksheet=tabox_drink_sheet[tabox_sheet_index])
            dr3 = pygsheets.datarange.DataRange(start="E2", end="E41", worksheet=tabox_drink_sheet[tabox_sheet_index])
            dr4 = pygsheets.datarange.DataRange(start="C2", end="C13", worksheet=tabox_drink_sheet[drink_sheet_index])
            dr5 = pygsheets.datarange.DataRange(start="D2", end="D13", worksheet=tabox_drink_sheet[drink_sheet_index])

            tabox_in_columns = tabox_df.iloc[:, 2].values.astype(int).tolist()
            tabox_out_columns = tabox_df.iloc[:, 3].values.astype(int).tolist()
            tabox_god_hand_columns = tabox_df.iloc[:, 4].values.astype(int).tolist()

            tabox_zeros_arr = np.zeros(box_num).astype(int).tolist()

            god_hand_reset = []

            for i in range(len(tabox_df)):
                if str(tabox_df.iloc[i, 1]).startswith("box"):
                    god_hand_reset += [0]
                else:
                    god_hand_reset += [-1]

            if tabox_in_columns != tabox_zeros_arr:
                in_update_matrix = []

                for z in tabox_zeros_arr:
                    in_update_matrix += [[z]]
            else:
                in_update_matrix = []

            if tabox_out_columns != tabox_zeros_arr:
                out_update_matrix = []

                for z in tabox_zeros_arr:
                    out_update_matrix += [[z]]
            else:
                out_update_matrix = []

            if tabox_god_hand_columns != god_hand_reset:
                god_hand_matrix = []

                for z in god_hand_reset:
                    god_hand_matrix += [[z]]
            else:
                god_hand_matrix = []


            if len(out_update_matrix) > 0 or len(in_update_matrix) > 0 or len(god_hand_matrix) > 0:
                if NOW < AUTO_TIME:
                    print("是否重置打包盒出入库归零?")
                    print()
                    action_req = option_num(["是", "否"])
                    time.sleep(0.25)
                    user_input = option_limit(action_req, input("在这里输入>>>: "))

                    if user_input == 0:
                        if len(in_update_matrix) > 0:
                            dr1.update_values(values = in_update_matrix)

                        else:
                            pass

                        if len(out_update_matrix) > 0:
                            dr2.update_values(values = out_update_matrix)

                        else:
                            pass

                        if len(god_hand_matrix) > 0:
                            dr3.update_values(values = god_hand_matrix)

                        else:
                            pass

                        tabox_df = tabox_drink_sheet[tabox_sheet_index].get_as_df()
                        print("打包盒出入库已重置归零")
                        print()


                    else:
                        tabox_df = tabox_drink_sheet[tabox_sheet_index].get_as_df()
                        print("好的，打包盒出入库不会重置归零")
                        print()
                else:
                    if len(in_update_matrix) > 0:
                            dr1.update_values(values = in_update_matrix)

                    else:
                        pass

                    if len(out_update_matrix) > 0:
                        dr2.update_values(values = out_update_matrix)

                    else:
                        pass

                    if len(god_hand_matrix) > 0:
                        dr3.update_values(values = god_hand_matrix)

                    else:
                        pass

                    tabox_df = tabox_drink_sheet[tabox_sheet_index].get_as_df()
                    print("打包盒出入库已自动重置归零")
                    print()

            drink_in_columns = drink_df.iloc[:,2].astype(int).values.tolist()
            drink_out_columns = drink_df.iloc[:,3].astype(int).values.tolist()


            drink_zeros_arr = np.zeros(drink_num).astype(int).tolist()

            if drink_in_columns != drink_zeros_arr:
                drink_in_matrix = []
                for z in drink_zeros_arr:
                    drink_in_matrix += [[z]]

            else:
                drink_in_matrix = []

            if drink_out_columns != drink_zeros_arr:
                drink_out_matrix = []
                for z in drink_zeros_arr:
                    drink_out_matrix += [[z]]
            else:
                drink_out_matrix = []

            if len(drink_in_matrix) > 0 or len(drink_out_matrix) > 0:
                if NOW < AUTO_TIME:
                    print("是否重置酒水出入库归零?")
                    print()
                    action_req = option_num(["是", "否"])
                    time.sleep(0.25)
                    user_input = option_limit(action_req, input("在这里输入>>>: "))

                    if user_input == 0:
                        if len(drink_in_matrix) > 0:
                            dr4.update_values(values = drink_in_matrix)

                        else:
                            pass

                        if len(drink_out_matrix) > 0:
                            dr5.update_values(values = drink_out_matrix)

                        else:
                            pass

                        drink_df = tabox_drink_sheet[drink_sheet_index].get_as_df()
                        print("酒水出入库已重置归零")
                        print()
                    else:
                        drink_df = tabox_drink_sheet[drink_sheet_index].get_as_df()
                        print("好的，酒水出入库不会重置归零")
                        print()
                else:
                    if len(drink_in_matrix) > 0:
                        dr4.update_values(values = drink_in_matrix)

                    else:
                        pass

                    if len(drink_out_matrix) > 0:
                        dr5.update_values(values = drink_out_matrix)

                    else:
                        pass

                    drink_df = tabox_drink_sheet[drink_sheet_index].get_as_df()
                    print("酒水出入库已自动重置归零")
                    print()

        except Exception as e:
            print(e)
            print()
            print("打包盒或酒水出入库可能重置归零失败！")
            print()

    else:
        tabox_df = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_db_filename), sheet_name=takeaway_box_sheetname)
        drink_df = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_db_filename), sheet_name=drink_stock_sheetname)

        tabox_names = []
        for index in range(box_num):
            tabox_names += [str(k_dict["box{}".format(index)])]

        tabox_df["BOX NAME"] = tabox_names
        tabox_df["IN (BY PIECE)"] = 0
        tabox_df["OUT (BY PIECE)"] = 0
        tabox_df["HAND OF GOD (BY PACK)"] = -1

        drink_names = []
        for index in range(drink_num):
            drink_names += [str(k_dict["drink{}".format(index)])]

        drink_df["DRINK NAME"] = drink_names
        drink_df["IN (BY PIECE)"] = 0
        drink_df["OUT (BY PIECE)"] = 0

    if wifi:
        rcv_df = get_rcv(fernet_key, k_dict, google_auth, backup_foldername, local_db_filename, True)
    else:
        rcv_df = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_db_filename), sheet_name=receivers_sheetname)

    #online database sheet
    if wifi:
        try:
            online_db_sheet = google_auth.open_by_url(online_database_url)

            #upload online database
            for name in sheet_names:
                df = take_databases[name]
                df["DATE"] = pd.to_datetime(df["DATE"])
                df["TIME LOG"] = pd.to_datetime(df["TIME LOG"])
                df["DATE"] = df["DATE"].apply(lambda y: y.strftime("%Y-%m-%d"))
                df["TIME LOG"] = df["TIME LOG"].apply(lambda x: x.strftime("%Y-%m-%d %H:%M:%S"))

                db_index = online_db_sheet.worksheet(property="title", value=name).index
                online_db_sheet[db_index].set_dataframe(df, start="A1", nan="")

            #print("网络上传数据库成功！")
        except Exception as e:
            print("网络上传失败，错误描述如下：")
            print(e)
            print()
    else:
        print()
        print("无网络连接，数据将会暂时本地保存")
        print("待网络连接后，本地暂时保存的数据将会自动上传到云端")

    if wifi:
        shift_db = retrieve_shift_db(online_database_url, shift_db_sheetname, local_db_filename, backup_foldername)
        shift_db["DATE"] = pd.to_datetime(shift_db["DATE"])
        shift_db["TIME LOG"] = pd.to_datetime(shift_db["TIME LOG"])
        shift_db["DATE"] = shift_db["DATE"].apply(lambda y : y.strftime("%Y-%m-%d"))
        shift_db["TIME LOG"] = shift_db["TIME LOG"].apply(lambda x : x.strftime("%Y-%m-%d %H:%M:%S"))
    else:
        shift_db = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_db_filename), sheet_name=shift_db_sheetname)
        shift_db["DATE"] = pd.to_datetime(shift_db["DATE"])
        shift_db["TIME LOG"] = pd.to_datetime(shift_db["TIME LOG"])
        shift_db["DATE"] = shift_db["DATE"].apply(lambda y : y.strftime("%Y-%m-%d"))
        shift_db["TIME LOG"] = shift_db["TIME LOG"].apply(lambda x : x.strftime("%Y-%m-%d %H:%M:%S"))

    with pd.ExcelWriter("{}/{}/{}".format(os.getcwd(), backup_foldername, local_db_filename)) as writer:
        for name in sheet_names:
            df = take_databases[name]
            df["DATE"] = pd.to_datetime(df["DATE"])
            df["TIME LOG"] = pd.to_datetime(df["TIME LOG"])
            df["DATE"] = df["DATE"].apply(lambda y: y.strftime("%Y-%m-%d"))
            df["TIME LOG"] = df["TIME LOG"].apply(lambda x: x.strftime("%Y-%m-%d %H:%M:%S"))

            df.to_excel(writer, sheet_name=name, index=False, header=True)

        shift_db.to_excel(writer, sheet_name=shift_db_sheetname, index=False, header=True)
        tabox_df.to_excel(writer, sheet_name=takeaway_box_sheetname, index=False, header=True)
        drink_df.to_excel(writer, sheet_name=drink_stock_sheetname, index=False, header=True)
        rcv_df.to_excel(writer, sheet_name=receivers_sheetname, index=False, header=True)

    #print("本地备份数据库成功！")

def sending_telegram(is_pr, message, api, receiver, wifi):
    if is_pr:
        print_string = ""
        for text in message:
            print_string += "{} \n".format(text)
    else:
        print_string = message

    if wifi:
        url = f"https://api.telegram.org/bot{api}/sendMessage?chat_id={receiver}&text={print_string}"
        print("正在给{}发送Telegram".format(receiver))
        resp = requests.get(url)

        if resp.status_code == 200:
            print("{}发送成功! ".format(receiver))

        else:
            print("{}发送失败! ".format(receiver))
    else:
        pass

def get_rcv(fernet_key, k_dict, google_auth, backup_foldername, local_database_filename, return_df):
    rcv_sheetname = k_dict["receivers_sheetname"]
    rcv_url = fernet_decrypt(k_dict["receivers_url"], fernet_key)

    try:
        rcv_sheet = google_auth.open_by_url(rcv_url)
        rcv_sheetname_index = rcv_sheet.worksheet(property="title", value=rcv_sheetname).index
        rcv_df = rcv_sheet[rcv_sheetname_index].get_as_df()
        is_encrypted = True

    except Exception as e:
        print()
        print("通过机器人读取收件人信息失败，错误描述如下：")
        print(e)
        print()
        rcv_df = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_database_filename), sheet_name=rcv_sheetname)
        is_encrypted = False
        print("已获取本地备份收件人信息")

    if return_df:
        if is_encrypted:
            rcv_df["ACCOUNT"] = rcv_df["ACCOUNT"].astype(str)
            rcv_df["ACCOUNT"] = rcv_df["ACCOUNT"].apply(lambda x : fernet_decrypt(x, fernet_key))
        else:
            pass

        return rcv_df

    else:
        rcv_email = {}
        rcv_telegram = {}
        for index in range(len(rcv_df)):
            if str(rcv_df.iloc[index, 1]).strip().upper() == "EMAIL":
                if is_encrypted:
                    rcv_email.update({ str(rcv_df.iloc[index, 0]).strip().upper() : fernet_decrypt(str(rcv_df.iloc[index, 2]), fernet_key) })
                else:
                    rcv_email.update({ str(rcv_df.iloc[index, 0]).strip().upper() : str(rcv_df.iloc[index, 2]) })

            elif str(rcv_df.iloc[index, 1]).strip().upper() == "TELEGRAM":
                if is_encrypted:
                    rcv_telegram.update({ str(rcv_df.iloc[index, 0]).strip().upper() : fernet_decrypt(str(rcv_df.iloc[index, 2]), fernet_key) })
                else:
                    rcv_telegram.update({ str(rcv_df.iloc[index, 0]).strip().upper() : str(rcv_df.iloc[index, 2]) })

            else:
                pass

        return rcv_telegram, rcv_email

def parse_sending(payslip_on_duty, drink_on_duty, box_on_duty, cashier_on_duty, google_auth, outlet, send_dict, drink_message_string, tabox_message_string, print_result, k_dict, fernet_key, wifi, date_dict, value_dict, db_writables, backup_foldername, database_url, manager_on_duty):
    database_url = fernet_decrypt(database_url, fernet_key)

    date = date_dict["dfb"].strftime("%Y-%m-%d")

    drink_stock_alert = eval(k_dict["drink_stock_alert"].strip().capitalize())
    box_stock_alert = eval(k_dict["box_stock_alert"].strip().capitalize())
    night_audit_alert = eval(k_dict["night_audit_alert"].strip().capitalize())
    payslip_end_month_alert = eval(k_dict["payslip_end_month_alert"].strip().capitalize())
    send_to_manager = eval(k_dict["send_to_manager"].strip().capitalize())

    send_drink_msg = send_dict["send_drink_msg"]
    send_tabox_msg = send_dict["send_tabox_msg"]
    send_print_result = send_dict["send_print_result"]

    night_audit_send_channel = k_dict["night_audit_send_channel"]
    drink_send_channel = k_dict["drink_send_channel"]
    tabox_send_channel = k_dict["tabox_send_channel"]
    payslip_send_channel = k_dict["payslip_time_send_channel"]

    shift_db_sheetname = k_dict["shift_database_sheetname"]
    local_database_filename = k_dict["local_database_filename"]

    cashier_on_duty = str(cashier_on_duty).strip().upper()
    drink_on_duty = str(drink_on_duty).strip().upper()
    box_on_duty = str(box_on_duty).strip().upper()
    payslip_on_duty = str(payslip_on_duty).strip().upper()
    manager_on_duty = str(manager_on_duty).strip().upper()

    write_finance_db = db_writables["write_finance_db"]

    outlet = str(outlet).strip().capitalize()

    rcv_telegram, rcv_email = get_rcv(fernet_key, k_dict, google_auth, backup_foldername, local_database_filename, False)

    print()
    print("警告: 哪怕接下来你已经收到了信息, 请确保该程序完全运行完成。")
    print("如果你已知晓，请按回车键继续运行程序。")
    time.sleep(0.25)
    input("在这里输入>>>:")
    print()
    print()

    if drink_stock_alert:

        drink_stock_email_server = fernet_decrypt(k_dict["drink_stock_email_server"], fernet_key)
        drink_stock_email_sender = fernet_decrypt(k_dict["drink_stock_email_sender"], fernet_key)
        drink_stock_sender_password = fernet_decrypt(k_dict["drink_stock_sender_password"], fernet_key)
        drink_stock_telegram_bot_api = fernet_decrypt(k_dict["drink_stock_telegram_bot_api"], fernet_key)

        if send_drink_msg:
            if len(drink_message_string) > 0:
                if drink_send_channel.strip().capitalize() == "Telegram":
                    drinkers = drink_on_duty.split(",")

                    for d in drinkers:
                        receivers = rcv_telegram[d]
                        sending_telegram(is_pr=False,
                                          message=drink_message_string,
                                          api = drink_stock_telegram_bot_api,
                                          receiver=receivers,
                                          wifi=wifi)

                elif drink_send_channel.strip().capitalize() == "Email":
                    drinkers = drink_on_duty.split(",")
                    for d in drinkers:
                        receivers = rcv_email[d]
                        sending_email(is_pr=False,
                                      mail_server=drink_stock_email_server,
                                      mail_sender=drink_stock_email_sender,
                                      mail_sender_password=drink_stock_sender_password,
                                      mail_receivers=receivers,
                                      mail_subject="{}的酒水库存通知{}".format(outlet, date),
                                      message_string = drink_message_string,
                                      wifi = wifi)
                else:
                    print("Drink alert sending channel is not defined correctly")
            else:
                pass
        else:
            pass
    else:
        pass

    if box_stock_alert:
        box_stock_email_server = fernet_decrypt(k_dict["box_stock_email_server"], fernet_key)
        box_stock_email_sender = fernet_decrypt(k_dict["box_stock_email_sender"], fernet_key)
        box_stock_sender_password = fernet_decrypt(k_dict["box_stock_sender_password"], fernet_key)
        box_stock_telegram_bot_api = fernet_decrypt(k_dict["box_stock_telegram_bot_api"], fernet_key)

        if send_tabox_msg:
            if len(tabox_message_string) > 0:
                if tabox_send_channel.strip().capitalize() == "Telegram":
                    boxers = box_on_duty.split(",")
                    for b in boxers:
                        receivers = rcv_telegram[b]
                        sending_telegram(is_pr=False,
                                          message=tabox_message_string,
                                          api = box_stock_telegram_bot_api,
                                          receiver=receivers,
                                          wifi=wifi)

                elif tabox_send_channel.strip().capitalize() == "Email":
                    boxers = box_on_duty.split(",")
                    for b in boxers:
                        receivers = rcv_email[b]
                        sending_email(is_pr=False,
                                      mail_server=box_stock_email_server,
                                      mail_sender=box_stock_email_sender,
                                      mail_sender_password=box_stock_sender_password,
                                      mail_receivers=receivers,
                                      mail_subject="{}的打包盒库存通知{}".format(outlet, date),
                                      message_string = tabox_message_string,
                                      wifi = wifi)
                else:
                    print("Takeaway box alert sending channel is not defined correctly")
            else:
                pass
        else:
            pass
    else:
        pass

    if night_audit_alert:
        night_audit_email_server = fernet_decrypt(k_dict["night_audit_email_server"], fernet_key)
        night_audit_email_sender = fernet_decrypt(k_dict["night_audit_email_sender"], fernet_key)
        night_audit_sender_password = fernet_decrypt(k_dict["night_audit_sender_password"], fernet_key)
        night_audit_telegram_bot_api = fernet_decrypt(k_dict["night_audit_telegram_bot_api"], fernet_key)

        svc = value_dict["svc"]
        gst = value_dict["gst"]
        ads = value_dict["ads"]

        if send_print_result:
            if len(print_result) > 0:
                if night_audit_send_channel.strip().capitalize() == "Telegram":
                    cashiers = cashier_on_duty.split(",")
                    for c in cashiers:
                        receivers = rcv_telegram[c]
                        sending_telegram(is_pr=True,
                                          message=print_result,
                                          api = night_audit_telegram_bot_api,
                                          receiver=receivers,
                                          wifi=wifi)

                        second_message = "服务费: ${}\n GST: ${}\n 日均营业额: ${}\n ".format(svc, gst, ads)
                        sending_telegram(is_pr=False,
                                         message=second_message,
                                         api = night_audit_telegram_bot_api,
                                         receiver=receivers,
                                         wifi=wifi)

                    if send_to_manager:
                        print()
                        print()
                        print("请查看Telegram, 信息可以直接发给经理吗? ")
                        time.sleep(0.25)
                        send_mgr_options = option_num(["信息准确无误,可以直接发给经理", "还有要修改的地方,我会自行发给经理"])
                        send_mgr_input = option_limit(send_mgr_options, input("在这里输入>>>: "))

                        if send_mgr_input == 0:
                            managers = manager_on_duty.split(",")
                            for m in managers:
                                receivers = rcv_telegram[m]
                                sending_telegram(is_pr=True,
                                                message=print_result,
                                                api = night_audit_telegram_bot_api,
                                                receiver=receivers,
                                                wifi=wifi)

                            reset_remarks = True

                        else:
                            print()
                            print("好的, 没有发给经理。")
                            reset_remarks = False

                    else:
                        reset_remarks = True

                elif night_audit_send_channel.strip().capitalize() == "Email":
                    print_result += ["——————————————", "服务费: ${}".format(svc),
                                     "GST: ${}".format(gst), "日均营业额: ${}".format(ads)]

                    cashiers = cashier_on_duty.split(",")
                    for c in cashiers:
                        receivers = rcv_email[c]
                        sending_email(is_pr=True,
                                      mail_server=night_audit_email_server,
                                      mail_sender=night_audit_email_sender,
                                      mail_sender_password=night_audit_sender_password,
                                      mail_receivers=receivers,
                                      mail_subject="{}的报表信息{}".format(outlet, date),
                                      message_string = print_result,
                                      wifi = wifi)

                    if send_to_manager:
                        print()
                        print()
                        print("请查看电邮, 信息可以直接发给经理吗? ")
                        time.sleep(0.25)
                        send_mgr_options = option_num(["信息准确无误,可以直接发给经理", "还有要修改的地方,我会自行发给经理"])
                        send_mgr_input = option_limit(send_mgr_options, input("在这里输入>>>: "))

                        if send_mgr_input == 0:
                            managers = manager_on_duty.split(",")
                            for m in managers:
                                receivers = rcv_email[m]
                                sending_email(is_pr=True,
                                            mail_server=night_audit_email_server,
                                            mail_sender=night_audit_email_sender,
                                            mail_sender_password=night_audit_sender_password,
                                            mail_receivers=receivers,
                                            mail_subject="{}的报表信息{}".format(outlet, date),
                                            message_string = print_result,
                                            wifi = wifi)

                            reset_remarks = True

                        else:
                            print()
                            print("好的, 没有发给经理。")
                            reset_remarks = False

                    else:
                        reset_remarks = True

                else:
                    print("Night audit alert sending channel is not defined correctly")
                    reset_remarks = False
        else:
            if wifi:
                if write_finance_db:
                    if len(print_result) > 0:

                        reset_remarks = False

                        print("是否重新发送关帐报表? ")
                        action_req = option_num(["重新发送", "不发送"])
                        time.sleep(0.25)
                        user_input = option_limit(action_req, input("在这里输入>>>: "))

                        if user_input == 0:
                            if night_audit_send_channel.strip().capitalize() == "Telegram":
                                cashiers = cashier_on_duty.split(",")
                                for c in cashiers:
                                    receivers = rcv_telegram[c]
                                    sending_telegram(is_pr=True,
                                                      message=print_result,
                                                      api = night_audit_telegram_bot_api,
                                                      receiver=receivers,
                                                      wifi=wifi)

                                    second_message = "服务费: ${}\n GST: ${}\n 日均营业额: ${}\n ".format(svc, gst, ads)
                                    sending_telegram(is_pr=False,
                                                     message=second_message,
                                                     api = night_audit_telegram_bot_api,
                                                     receiver=receivers,
                                                     wifi=wifi)

                                if send_to_manager:
                                    print()
                                    print()
                                    print("请查看Telegram, 信息重新发给经理吗? ")
                                    time.sleep(0.25)
                                    send_mgr_options = option_num(["重新发送", "不重新发送"])
                                    send_mgr_input = option_limit(send_mgr_options, input("在这里输入>>>: "))

                                    if send_mgr_input == 0:
                                        managers = manager_on_duty.split(",")
                                        for m in managers:
                                            receivers = rcv_telegram[m]
                                            sending_telegram(is_pr=True,
                                                            message=print_result,
                                                            api = night_audit_telegram_bot_api,
                                                            receiver=receivers,
                                                            wifi=wifi)

                                        reset_remarks = False
                                    else:
                                        print()
                                        print("好的, 没有重新发给经理。")
                                        reset_remarks = False


                            elif night_audit_send_channel.strip().capitalize() == "Email":
                                print_result += ["——————————————", "服务费: ${}".format(svc),
                                                 "GST: ${}".format(gst), "日均营业额: ${}".format(ads)]

                                cashiers = cashier_on_duty.split(",")

                                for c in cashiers:
                                    receivers = rcv_email[c]

                                    sending_email(is_pr=True,
                                                  mail_server=night_audit_email_server,
                                                  mail_sender=night_audit_email_sender,
                                                  mail_sender_password=night_audit_sender_password,
                                                  mail_receivers=receivers,
                                                  mail_subject="{}的报表信息{}".format(outlet, date),
                                                  message_string = print_result,
                                                  wifi = wifi)

                                if send_to_manager:
                                    print()
                                    print()
                                    print("请查看电邮, 信息重新发给经理吗? ")
                                    time.sleep(0.25)
                                    send_mgr_options = option_num(["重新发送", "不重新发送"])
                                    send_mgr_input = option_limit(send_mgr_options, input("在这里输入>>>: "))

                                    if send_mgr_input == 0:
                                        managers = manager_on_duty.split(",")
                                        for m in managers:
                                            receivers = rcv_email[m]
                                            sending_email(is_pr=True,
                                                        mail_server=night_audit_email_server,
                                                        mail_sender=night_audit_email_sender,
                                                        mail_sender_password=night_audit_sender_password,
                                                        mail_receivers=receivers,
                                                        mail_subject="{}的报表信息{}".format(outlet, date),
                                                        message_string = print_result,
                                                        wifi = wifi)

                                        reset_remarks = False
                                    else:
                                        print()
                                        print("好的, 没有重新发给经理。")
                                        reset_remarks = False

                            else:
                                print("Night audit alert sending channel is not defined correctly")

        feedback_reset(k_dict, reset_remarks)

    else:
        pass

    if payslip_end_month_alert:
        date_from_book = pd.to_datetime(date_dict["dfb"])
        LAST_DAY = month_last_day(year=date_from_book.year, month=date_from_book.month)
        LAST_DAY_OF_MONTH = pd.to_datetime(dt.datetime(year=date_from_book.year, month=date_from_book.month, day=LAST_DAY))

        if date_from_book == LAST_DAY_OF_MONTH:
            shift_db = retrieve_shift_db(database_url, shift_db_sheetname, local_database_filename, backup_foldername)

            if isinstance(shift_db, pd.DataFrame):
                shift_db["DATE"] = pd.to_datetime(shift_db["DATE"])
                shift_db["ID"] = shift_db["ID"].astype(int)
                shift_db["ID"] = shift_db["ID"].astype(str)

                shiftURL = fernet_decrypt(k_dict["shift_url"], fernet_key)

                #loading employee_info
                df = google_auth.open_by_url(shiftURL)
                df = df[2].get_as_df()
                df = df[df.iloc[:,0].astype(str).str.len() > 0]
                df["ID"] = df["ID"].astype(int)
                df["ID"] = df["ID"].astype(str)
                df["FIRST DAY DATE"] = pd.to_datetime(df["FIRST DAY DATE"])
                df["AL START DATE"] = pd.to_datetime(df["AL START DATE"])
                df["AL END DATE"] = pd.to_datetime(df["AL END DATE"])

                start_date = dt.datetime(date_from_book.year, date_from_book.month, 1)
                end_date = LAST_DAY_OF_MONTH

                payslip_time_df = payslip_time(shift_database = shift_db,
                                               start_date = pd.to_datetime(start_date),
                                               end_date = pd.to_datetime(end_date),
                                               employee_info_df=df)

                payslip_time_df.reset_index(inplace=True)
                payslip_time_df.drop("index", axis=1, inplace=True)

                statement_dict = payslip_time_statement(concat_df = payslip_time_df,
                                                        employee_info_df = df,
                                                        start_date = pd.to_datetime(start_date),
                                                        end_date = pd.to_datetime(end_date) )


                send_string = "{}年{}月{}日至{}年{}月{}日{}员工工资工时分析单 \n ".format(start_date.year, start_date.month, start_date.day, end_date.year, end_date.month, end_date.day, outlet)
                send_string += " \n "

                for k, i in statement_dict.items():
                    send_string += "{} : {} \n ".format(k, i)

                if send_print_result:
                    if payslip_send_channel.strip().capitalize() == "Telegram":
                        payslip_rcv = rcv_telegram[payslip_on_duty]
                        telegram_api = fernet_decrypt(k_dict["payslip_time_telegram_bot_api"], fernet_key)

                        sending_telegram(is_pr=False,
                                         message=send_string,
                                         api = telegram_api,
                                         receiver=payslip_rcv,
                                         wifi=wifi)

                    elif payslip_send_channel.strip().capitalize() == "Email":
                        email_server = fernet_decrypt(k_dict["payslip_time_email_server"], fernet_key)
                        email_sender = fernet_decrypt(k_dict["payslip_time_email_sender"], fernet_key)
                        email_sender_password = fernet_decrypt(k_dict["payslip_time_sender_password"], fernet_key)
                        email_receiver = rcv_email[payslip_on_duty]

                        sending_email(is_pr=False,
                                      mail_server=email_server,
                                      mail_sender=email_sender,
                                      mail_sender_password=email_sender_password,
                                      mail_receivers=email_receiver,
                                      mail_subject="{}的工时分析".format(outlet),
                                      message_string=send_string,
                                      wifi=wifi)

                    else:
                        print("Payslip time send channel is not defined correctly")


                else:
                    if wifi:
                        if write_finance_db:
                            print("是否重新发送工时工资分析单? ")
                            action_req = option_num(["重新发送", "不发送"])
                            time.sleep(0.25)
                            user_input = option_limit(action_req, input("在这里输入>>>: "))

                            if user_input == 0:
                                if payslip_send_channel.strip().capitalize() == "Telegram":
                                    payslip_rcv = rcv_telegram[payslip_on_duty]
                                    telegram_api = fernet_decrypt(k_dict["payslip_time_telegram_bot_api"], fernet_key)

                                    sending_telegram(is_pr=False,
                                                     message=send_string,
                                                     api = telegram_api,
                                                     receiver=payslip_rcv,
                                                     wifi=wifi)

                                elif payslip_send_channel.strip().capitalize() == "Email":
                                    email_server = fernet_decrypt(k_dict["payslip_time_email_server"], fernet_key)
                                    email_sender = fernet_decrypt(k_dict["payslip_time_email_sender"], fernet_key)
                                    email_sender_password = fernet_decrypt(k_dict["payslip_time_sender_password"], fernet_key)
                                    email_receiver = rcv_email[payslip_on_duty]

                                    sending_email(is_pr=False,
                                                  mail_server=email_server,
                                                  mail_sender=email_sender,
                                                  mail_sender_password=email_sender_password,
                                                  mail_receivers=email_receiver,
                                                  mail_subject="{}的工时分析".format(outlet),
                                                  message_string=send_string,
                                                  wifi=wifi)

                                else:
                                    print("Payslip time send channel is not defined correctly")

                            else:
                                pass
                        else:
                            pass
                    else:
                        pass

            else:
                print("无法获取排班数据库")
        else:
            pass
    else:
        pass

def parse_display_df(value_dict, rule_df_dict, k_dict, google_auth, db_writables, fernet_key, database_url, backup_foldername):
    tabox_inv_function = eval(k_dict["takeaway_box_inventory"].strip().capitalize())
    drink_inv_function = eval(k_dict["drink_inventory"].strip().capitalize())

    bookkeeping_columns = ["描述", "收入", "支出", "月累计"]
    drink_columns = ["饮料(单位)", "入库", "出库", "剩余"]
    tabox_columns = ["物品(单位)", "入库", "用出", "剩余"]

    write_finance_db = db_writables["write_finance_db"]
    tabox_write_db = db_writables["tabox_write_db"]
    write_drink_db = db_writables["write_drink_db"]

    box_drink_in_out_url = fernet_decrypt(k_dict["box_drink_in_out_url"], fernet_key)
    tabox_sheetname = k_dict["takeaway_box_sheetname"]
    drink_stock_sheetname = k_dict["drink_stock_sheetname"]

    local_db_filename = k_dict["local_database_filename"]

    database_url = fernet_decrypt(database_url, fernet_key)

    takeaway_box_database_sheetname = k_dict["takeaway_box_database_sheetname"]
    drink_db_sheetname = k_dict["drink_database_sheetname"]

    vd_df = rule_df_dict["vd_df"]

    if write_finance_db:
        original_balance = float(value_dict["cmns_ytd"]) + float(value_dict["net_sales"])

        particulars = ["初始额", "销售额"]
        expense = [0, 0]
        income = [0, float(value_dict["net_sales"])]
        balance = [float(value_dict["cmns_ytd"]),original_balance]


        for i in range(len(vd_df)):
            if int(vd_df.iloc[i, 6]) == 1:
                particulars += [str(vd_df.iloc[i, 7])]
                expense += [float(value_dict[str(vd_df.iloc[i,0])])]
                original_balance -= float(value_dict[str(vd_df.iloc[i,0])])
                balance += [original_balance]
                income += [0]
            else:
                continue

        particulars += ["今日关账合计"]
        expense_sum = sum(expense)
        expense += [expense_sum]
        income_sum = sum(income)
        income += [income_sum]
        balance += [original_balance]

        all_list = [particulars, income, expense, balance]

        bk_df = {}
        for index in range(len(bookkeeping_columns)):
            bk_df.update({ bookkeeping_columns[index] : all_list[index] })

        bk_df = pd.DataFrame(bk_df)

        dropIndex = []
        for i in range(len(bk_df)):
            if i == 0:
                continue
            else:
                if bk_df.iloc[i, 1] == 0 and bk_df.iloc[i, 2] == 0:
                    dropIndex += [i]
                else:
                    continue

        bk_df.drop(dropIndex, axis=0, inplace=True)
        bk_df.reset_index(inplace=True)
        bk_df.drop("index", axis=1, inplace=True)
        bk_df.replace(to_replace=0, value="", inplace=True)
    else:
        bk_df = pd.DataFrame(columns=bookkeeping_columns)

    if tabox_inv_function:
        if tabox_write_db:
            try:
                tabox_drink_sheet = google_auth.open_by_url(box_drink_in_out_url)
                tabox_sheetname_index = tabox_drink_sheet.worksheet(property="title", value=tabox_sheetname).index
                tabox_df = tabox_drink_sheet[tabox_sheetname_index].get_as_df()

            except:
                tabox_df = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_db_filename), sheet_name=tabox_sheetname)


            try:
                tabox_db_sheet = google_auth.open_by_url(database_url)
                tabox_db_sheetname_index = tabox_db_sheet.worksheet(property="title", value=takeaway_box_database_sheetname).index
                tabox_db = tabox_db_sheet[tabox_db_sheetname_index].get_as_df()

            except:
                tabox_db = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_db_filename), sheet_name = takeaway_box_database_sheetname)

            tabox_db["DATE"] = pd.to_datetime(tabox_db["DATE"])
            tabox_max_date = tabox_db["DATE"].max()
            tabox_db = tabox_db[tabox_db["DATE"] == tabox_max_date]

            names = []
            box_in = []
            box_out = []
            box_value_now = []
            for index in range(len(tabox_df)):
                NAME = str(tabox_df.iloc[index, 9])

                if NAME.startswith("box"):
                    continue
                else:
                    IN = float(tabox_db["{}入库".format(tabox_df.iloc[index,1])].values[0])
                    OUT = float(tabox_db["{}出库".format(tabox_df.iloc[index, 1])].values[0])
                    NOW = float(tabox_db["{}现有数量".format(tabox_df.iloc[index, 1])].values[0])

                    names += [NAME]
                    box_in += [IN]
                    box_out += [OUT]
                    box_value_now += [NOW]

            all_list = [names, box_in, box_out, box_value_now]
            show_box_df = {}
            for i in range(len(tabox_columns)):
                show_box_df.update({ tabox_columns[i] : all_list[i] })

            show_box_df = pd.DataFrame(show_box_df)

        else:
            show_box_df = pd.DataFrame(columns=tabox_columns)
            tabox_max_date = dt.datetime.now()
    else:
        show_box_df = pd.DataFrame(columns=tabox_columns)
        tabox_max_date = dt.datetime.now()

    if drink_inv_function:
        if write_drink_db:
            try:
                tabox_drink_sheet = google_auth.open_by_url(box_drink_in_out_url)
                drink_sheetname_index = tabox_drink_sheet.worksheet(property="title", value=drink_stock_sheetname).index
                drink_df = tabox_drink_sheet[drink_sheetname_index].get_as_df()

            except:
                drink_df = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_db_filename), sheet_name=drink_stock_sheetname)

            try:
                drink_db_sheet = google_auth.open_by_url(database_url)
                drink_db_sheetname_index = drink_db_sheet.worksheet(property="title", value=drink_db_sheetname).index
                drink_db = drink_db_sheet[drink_db_sheetname_index].get_as_df()
            except:
                drink_db = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_db_filename), sheet_name=drink_db_sheetname)

            drink_db["DATE"] = pd.to_datetime(drink_db["DATE"])
            drink_db_max_date = drink_db["DATE"].max()
            drink_db = drink_db[drink_db["DATE"] == drink_db_max_date]

            name = []
            drink_in = []
            drink_out = []
            drink_left = []

            for index in range(len(drink_df)):
                NAME = drink_df.iloc[index, 6]

                if NAME.startswith("饮料"):
                    continue

                else:
                    IN = int(drink_db["{}进".format(drink_df.iloc[index,1])].values[0])
                    OUT = int(drink_db["{}出".format(drink_df.iloc[index, 1])].values[0])
                    LEFT = int(drink_db["{}实结存".format(drink_df.iloc[index,1])].values[0])

                    name += [NAME]
                    drink_in += [IN]
                    drink_out += [OUT]
                    drink_left += [LEFT]

            all_list = [name, drink_in, drink_out, drink_left]

            show_drink_df = {}
            for i in range(len(drink_columns)):
                show_drink_df.update({ drink_columns[i] : all_list[i] })

            show_drink_df = pd.DataFrame(show_drink_df)
        else:
            show_drink_df = pd.DataFrame(columns=drink_columns)
            drink_db_max_date = dt.datetime.now()
    else:
        show_drink_df = pd.DataFrame(columns=drink_columns)
        drink_db_max_date = dt.datetime.now()

    return bk_df, show_box_df, show_drink_df, tabox_max_date, drink_db_max_date

def prtdf(df):
    with pd.option_context('display.max_rows', None,
                           'display.max_columns', None,
                           'display.width', 1000,
                           'display.precision', 2,
                           'display.colheader_justify', 'center'):
        return display(df)

def backup_script(script_backup_filename, script, backup_foldername):
    with open("fernet_key.txt", "rb") as keyfile:
        fernet_key = keyfile.read()

    f_handler = Fernet(fernet_key)
    encrypted_script = f_handler.encrypt(script.encode())

    if not os.path.exists("{}/{}".format(os.getcwd(), backup_foldername)):
        os.makedirs("{}/{}".format(os.getcwd(), backup_foldername))

    with open("{}/{}/{}".format(os.getcwd(), backup_foldername, script_backup_filename), "wb") as sfile:
        sfile.write(encrypted_script)

def night_audit_main(database_url, db_setting_url, serialized_rule_filename, service_filename, constants_sheetname, google_auth, box_num, drink_num, promo_num, lun_sales, lun_gc, tb_sales, tb_gc, lun_fwc, lun_kwc, tb_fwc, tb_kwc, night_fwc, night_kwc, script_backup_filename, script, wifi, backup_foldername,cashier_on_duty, drink_on_duty, box_on_duty, payslip_on_duty, manager_on_duty):
    database_url = database_url
    db_setting_url = db_setting_url
    serialized_rule_filename = serialized_rule_filename
    constants_sheetname = constants_sheetname
    service_filename = service_filename
    google_auth = google_auth
    box_num = box_num
    drink_num = drink_num
    promo_num = promo_num
    lun_sales = lun_sales
    lun_gc = lun_gc
    lun_fwc = lun_fwc
    lun_kwc = lun_kwc
    tb_sales = tb_sales
    tb_gc = tb_gc
    tb_fwc = tb_fwc
    tb_kwc = tb_kwc
    night_fwc = night_fwc
    night_kwc = night_kwc
    script_backup_filename = script_backup_filename
    script = script
    wifi = wifi
    backup_foldername = backup_foldername
    cashier_on_duty = cashier_on_duty
    drink_on_duty = drink_on_duty
    box_on_duty = box_on_duty
    payslip_on_duty = payslip_on_duty
    manager_on_duty = manager_on_duty

    fernet_key = get_key()

    if fernet_key == 0:
        print("安全密钥错误! ")
    else:
        res = pyfiglet.figlet_format("Night Audit")
        print(res)
        time.sleep(0.15)

        backup_script(script_backup_filename, script, backup_foldername)

        with tqdm(total=100) as pbar:
            if wifi:
                pbar.set_description("云端模式")
            else:
                pbar.set_description("离线模式")

            outlet = get_outlet()
            k_dict, rule_df_dict = get_dfs(google_auth, db_setting_url, constants_sheetname, serialized_rule_filename, outlet, fernet_key, wifi, backup_foldername)
            book_dict = get_book_dfs(k_dict, outlet)

            if len(book_dict) == 0:
                pass
            else:
                pbar.set_description("获取日期")
                date_dict = get_df_date(book_dict)
                pbar.update(3)

                pbar.set_description("获取列名称")
                get_columns = get_db_columns(k_dict, drink_num, box_num, promo_num)
                pbar.update(3)

                pbar.set_description("获取数据库")
                take_databases = retrieve_database(k_dict, google_auth, fernet_key, wifi, database_url, backup_foldername)
                pbar.update(3)

                pbar.set_description("更新列名称")
                take_databases = update_columns(take_databases, get_columns, k_dict)
                pbar.update(3)

                pbar.set_description("处理打包盒库存")
                tabox_write_db, box_value, box_unparsed_alert = parse_tabox(k_dict, google_auth, fernet_key, box_num, rule_df_dict, take_databases, book_dict, date_dict, wifi, backup_foldername)
                pbar.update(10)

                pbar.set_description("处理打包盒警示信息")
                box_stock_alert = eval(str(k_dict["box_stock_alert"]).strip().capitalize())
                box_alert_freq = str(k_dict["box_stock_alert_frequency"]).strip().upper()
                box_understock_alert = eval(str(k_dict["box_understocking_alert"]).strip().capitalize())
                box_overstock_alert = eval(str(k_dict["box_overstocking_alert"]).strip().capitalize())

                tabox_message_string = parse_alert(stock_alert_bool=box_stock_alert,
                                                 date_dict=date_dict,
                                                 unparsed_alert=box_unparsed_alert,
                                                 understock_alert_bool=box_understock_alert,
                                                 overstock_alert_bool=box_overstock_alert,
                                                 alert_freq=box_alert_freq,
                                                 alert_type="box",
                                                 outlet=outlet)

                pbar.update(10)

                pbar.set_description("计算营业额")
                value_dict, write_finance_db, write_promo_db = parse_value_dict(promo_num, lun_sales, lun_gc, tb_sales, tb_gc, lun_fwc, lun_kwc, tb_fwc, tb_kwc, night_fwc, night_kwc, rule_df_dict, take_databases, date_dict, book_dict, k_dict)
                pbar.update(4)

                pbar.set_description("处理酒水库存")
                write_drink_db, drink_dict, drink_unparsed_alert = parse_drink_stock(k_dict, fernet_key, google_auth, drink_num, value_dict, take_databases, date_dict, wifi, backup_foldername)
                pbar.update(4)

                pbar.set_description("处理酒水警示信息")
                drink_stock_alert = eval(str(k_dict["drink_stock_alert"]).strip().capitalize())
                pbar.update(4)

                drink_understock_alert = eval(str(k_dict["drink_understocking_alert"]).strip().capitalize())
                pbar.update(6)

                drink_overstock_alert = eval(str(k_dict["drink_overstocking_alert"]).strip().capitalize())
                pbar.update(4)

                drink_alert_freq = str(k_dict["drink_stock_alert_frequency"]).strip().upper()
                pbar.update(4)

                drink_message_string = parse_alert(stock_alert_bool=drink_stock_alert,
                                                   date_dict=date_dict,
                                                   unparsed_alert=drink_unparsed_alert,
                                                   understock_alert_bool=drink_understock_alert,
                                                   overstock_alert_bool=drink_overstock_alert,
                                                   alert_freq=drink_alert_freq,
                                                   alert_type="drink",
                                                   outlet=outlet)

                pbar.update(12)

                pbar.set_description("整理报表")
                print_result = parse_print_rule(value_dict, rule_df_dict, write_finance_db, write_promo_db, write_drink_db, tabox_write_db, k_dict)
                pbar.update(5)

                pbar.set_description("更新数据库")
                db_writables = db_write(write_drink_db, tabox_write_db, write_finance_db, write_promo_db)
                pbar.update(5)

                take_databases, send_dict = pending_upload_db(take_databases, k_dict, box_value, drink_dict, value_dict, db_writables, date_dict, drink_num, promo_num, get_columns)
                pbar.update(5)

                pbar.set_description("上传数据库")
                upload_db(database_url, take_databases, k_dict, fernet_key, google_auth, box_num, drink_num, wifi, backup_foldername)
                pbar.update(5)

                pbar.set_description("发信息")

                try:
                    parse_sending(payslip_on_duty, drink_on_duty, box_on_duty, cashier_on_duty, google_auth, outlet, send_dict, drink_message_string, tabox_message_string, print_result, k_dict, fernet_key, wifi, date_dict, value_dict, db_writables, backup_foldername, database_url, manager_on_duty)

                except Exception as e:
                    print("发送信息失败，错误描述如下: ")
                    print(e)
                    print()
                    print("已跳过发送信息")
                    print()
                    print()

                pbar.update(5)

                pbar.set_description("生成表格")
                bk_df, show_box_df, show_drink_df, tabox_max_date, drink_db_max_date = parse_display_df(value_dict, rule_df_dict, k_dict, google_auth, db_writables, fernet_key, database_url, backup_foldername)

                pbar.set_description("酒水明细")
                trigger_drink_generate = trigger_gen_drink_pdf(k_dict, date_dict)
                if trigger_drink_generate:
                    stock_count_foldername = "{}盘点文件".format(outlet.strip().capitalize())
                    songti_filename = "SongTi.ttf"

                    if not os.path.exists("{}/{}/{}".format(os.getcwd(), stock_count_foldername, songti_filename)):
                        print("宋体TTF文件'{}'不存在, 酒水明细表的生成无法继续".format(songti_filename))
                        print("请把宋体TTF文件'{}'保存至盘点文件名的目录下, 具体路径需在:'{}/{}/{}'".format(songti_filename, os.getcwd(), stock_count_foldername, songti_filename))
                        print()
                        print("生成酒水明细表PDF环节已跳过! ")

                    else:
                        print("月底了，已自动进入酒水明细表生成环节! ")
                        print()
                        print("读取字体文件中...")
                        custom_font_path = pathlib.Path("{}/{}/{}".format(os.getcwd(), stock_count_foldername, songti_filename))
                        borb_custom_font = borb_TrueTypeFont.true_type_font_from_file(custom_font_path)
                        print("字体文件读取完成。")
                        print()
                        print("进入酒水明细表生成环节后请准确选择你要生成的年份和月份")
                        print()
                        gen_drink_pdf(google_auth, db_setting_url, constants_sheetname, serialized_rule_filename, backup_foldername, database_url, borb_custom_font, drink_num)
                else:
                    pass

                pbar.set_description("任务完成")
                pbar.update(5)
                time.sleep(0.5)

                print()
                print("——————————————")
                for line in print_result:
                    print(line)

                print()
                print()
                print("——————————————")
                print()
                print("服务费: ${}".format(value_dict["svc"]))
                print("GST: ${}".format(value_dict["gst"]))
                print("日均营业额: ${}".format(value_dict["ads"]))
                print()
                if not wifi:
                    print("——————————————")
                    print()
                    if len(tabox_message_string) > 0:
                        print(tabox_message_string)
                        print()

                    if len(drink_message_string) > 0:
                        print(drink_message_string)
                print("——————————————")
                if not bk_df.empty:
                    print("账务簿记: ")
                    print(date_dict["date_chinese_string"])
                    prtdf(bk_df)
                    print()

                if not show_drink_df.empty:
                    print("酒水信息: ")
                    print(drink_db_max_date.strftime("%Y-%m-%d"))
                    prtdf(show_drink_df)
                    print()

                if not show_box_df.empty:
                    print("打包盒信息: ")
                    print(tabox_max_date.strftime("%Y-%m-%d"))
                    prtdf(show_box_df)
                    print()

def get_k_dictionary(google_auth, db_setting_url, constants_sheetname, serialized_rule_filename, outlet, fernet_key, backup_foldername):
    try:
        db_setting_url = fernet_decrypt(db_setting_url, fernet_key)

        db_setting_sheet = google_auth.open_by_url(db_setting_url)

        #k_df
        k_sheet_index = db_setting_sheet.worksheet(property="title", value=constants_sheetname).index
        k_df = db_setting_sheet[k_sheet_index].get_as_df()

        outlet_col_index = k_df.columns.get_loc(key=str(outlet).strip().upper())

        k_dict = {}
        for index in range(len(k_df)):
            k_dict.update({ str(k_df.iloc[index, 0]) : str(k_df.iloc[index, outlet_col_index]) })

    except Exception as e:
        print()
        print("从网络读取规则文件失败，错误描述如下: ")
        print()
        print(e)

        with open("{}/{}/{}".format(os.getcwd(), backup_foldername, serialized_rule_filename), "rb") as f:
            backup_list = pickle.load(f)

        k_df = backup_list[0]

        outlet_col_index = k_df.columns.get_loc(key=outlet)

        k_dict = {}
        for index in range(len(k_df)):
            k_dict.update({ str(k_df.iloc[index, 0]) : str(k_df.iloc[index, outlet_col_index]) })

        print("已采用本地备份规则文件数据")

    return k_dict

def get_inv_df(formURL, df_name):
    drop_list_dict = {
        "breakages_df":['选择操作','行为',
                        '你存入了什么物品？', '你把该物品存到了哪里？', '你存入了多少个该物品？', '你拿出了什么物品？', '你从哪里拿出了该物品？',
                        '你拿出了多少个该物品？', '你移动了什么物品？', '你移动了多少该物品？', '你从哪里移动了该物品？', '你要把该物品移到哪里？',
                        '选择进货物品', '进货数量', '盘点日期', '啤酒杯(个)_盘点数量', '红酒杯(个)_盘点数量', '饮料杯(个)_盘点数量',
                        '条纹杯(个)_盘点数量', '水杯(个)_盘点数量', '甜品碗(个)_盘点数量', '醋罐(个)_盘点数量', '汤碗(个)_盘点数量',
                        '茶杯(个)_盘点数量', '酱碟(个)_盘点数量', '短瓷汤匙(个)_盘点数量', '三人行水杯(个)_盘点数量',
                        '元宝碗(个)_盘点数量', '筷架(个)_盘点数量', '茶壶(个)_盘点数量', '骨碟(个)_盘点数量', '大分羹(个)_盘点数量',
                        '中分羹(个)_盘点数量', '不锈钢刀(个)_盘点数量', '不锈钢叉(个)_盘点数量', '焖面勺(个)_盘点数量',
                        '甜品叉(个)_盘点数量', '短甜品汤匙(个)_盘点数量', '长甜品汤匙(个)_盘点数量', '蟹钳(个)_盘点数量',
                        '保温水壶(个)_盘点数量', '长塑料汤匙(个)_盘点数量', '筷子(双)_盘点数量', '儿童碗(个)_盘点数量',
                        '儿童杯(个)_盘点数量', '儿童叉(个)_盘点数量', '儿童汤匙(个)_盘点数量', '酱清罐(个)_盘点数量',
                        '辣椒罐(个)_盘点数量', '牙签罐(个)_盘点数量', '胡椒粉罐(个)_盘点数量', '椒盐罐(个)_盘点数量',
                        '冷水壶(个)_盘点数量', '洗手盅(个)_盘点数量', '小冰桶(个)_盘点数量', '圆托盘(个)_盘点数量',
                        '长方托盘(个)_盘点数量', '娃娃椅(把)_盘点数量', 'MS 4(条)_盘点数量', 'JP 25ml(条)_盘点数量',
                        'SW203(条)_盘点数量', '黑色吸管(包)_盘点数量', '5*8线袋(包)_盘点数量', '7*9线袋(包)_盘点数量',
                        '小透明塑料袋(扎)_盘点数量', '中透明塑料袋(扎)_盘点数量', '一次性筷子(包)_盘点数量', '一次性汤匙(包)_盘点数量',
                        '短30*300保鲜膜(条)_盘点数量', '长45*300保鲜膜(条)_盘点数量', '0.05*36*48垃圾袋(小包)_盘点数量',
                        '0.035*30*34垃圾袋(小包)_盘点数量', '牙签(盒)_盘点数量', '菊花茶(公斤)_盘点数量', '普洱茶(公斤)_盘点数量',
                        '干抽纸(小包)_盘点数量', '吸油纸(卷)_盘点数量', '妈妈柠檬(桶)_盘点数量', '蓝色消毒液(桶)_盘点数量',
                        '拖地消毒液(桶)_盘点数量', 'TP200(条)_盘点数量', '大热敏纸(条)_盘点数量', '小热敏纸(条)_盘点数量'],

        "buyInStockDf":['选择操作', '选择破损物品', '破损数量', '你是谁?', '行为',
                        '你存入了什么物品？', '你把该物品存到了哪里？', '你存入了多少个该物品？', '你拿出了什么物品？', '你从哪里拿出了该物品？',
                        '你拿出了多少个该物品？', '你移动了什么物品？', '你移动了多少该物品？', '你从哪里移动了该物品？', '你要把该物品移到哪里？',
                        '盘点日期', '啤酒杯(个)_盘点数量', '红酒杯(个)_盘点数量', '饮料杯(个)_盘点数量',
                        '条纹杯(个)_盘点数量', '水杯(个)_盘点数量', '甜品碗(个)_盘点数量', '醋罐(个)_盘点数量', '汤碗(个)_盘点数量',
                        '茶杯(个)_盘点数量', '酱碟(个)_盘点数量', '短瓷汤匙(个)_盘点数量', '三人行水杯(个)_盘点数量',
                        '元宝碗(个)_盘点数量', '筷架(个)_盘点数量', '茶壶(个)_盘点数量', '骨碟(个)_盘点数量', '大分羹(个)_盘点数量',
                        '中分羹(个)_盘点数量', '不锈钢刀(个)_盘点数量', '不锈钢叉(个)_盘点数量', '焖面勺(个)_盘点数量',
                        '甜品叉(个)_盘点数量', '短甜品汤匙(个)_盘点数量', '长甜品汤匙(个)_盘点数量', '蟹钳(个)_盘点数量',
                        '保温水壶(个)_盘点数量', '长塑料汤匙(个)_盘点数量', '筷子(双)_盘点数量', '儿童碗(个)_盘点数量',
                        '儿童杯(个)_盘点数量', '儿童叉(个)_盘点数量', '儿童汤匙(个)_盘点数量', '酱清罐(个)_盘点数量',
                        '辣椒罐(个)_盘点数量', '牙签罐(个)_盘点数量', '胡椒粉罐(个)_盘点数量', '椒盐罐(个)_盘点数量',
                        '冷水壶(个)_盘点数量', '洗手盅(个)_盘点数量', '小冰桶(个)_盘点数量', '圆托盘(个)_盘点数量',
                        '长方托盘(个)_盘点数量', '娃娃椅(把)_盘点数量', 'MS 4(条)_盘点数量', 'JP 25ml(条)_盘点数量',
                        'SW203(条)_盘点数量', '黑色吸管(包)_盘点数量', '5*8线袋(包)_盘点数量', '7*9线袋(包)_盘点数量',
                        '小透明塑料袋(扎)_盘点数量', '中透明塑料袋(扎)_盘点数量', '一次性筷子(包)_盘点数量', '一次性汤匙(包)_盘点数量',
                        '短30*300保鲜膜(条)_盘点数量', '长45*300保鲜膜(条)_盘点数量', '0.05*36*48垃圾袋(小包)_盘点数量',
                        '0.035*30*34垃圾袋(小包)_盘点数量', '牙签(盒)_盘点数量', '菊花茶(公斤)_盘点数量', '普洱茶(公斤)_盘点数量',
                        '干抽纸(小包)_盘点数量', '吸油纸(卷)_盘点数量', '妈妈柠檬(桶)_盘点数量', '蓝色消毒液(桶)_盘点数量',
                        '拖地消毒液(桶)_盘点数量', 'TP200(条)_盘点数量', '大热敏纸(条)_盘点数量', '小热敏纸(条)_盘点数量'],
    }

    column_list_dict = {
        "breakages_df": ["Timestamp", "日期(年年年年-月月-日日)", "破损物品(单位)", "破损数量", "录入员"],
        "buyInStockDf": ['Timestamp', '进货日期(年年年年-月月-日日)', '进货物品(单位)', '进货数量'],
    }

    index_list_dict = {
        "breakages_df": 0,
        "buyInStockDf": 0,
        "lockInStockDf": 1,
        "stockCountDf": 0,
        "pageOneUnitPrice": 2,
        "pageTwoUnitPrice": 3,
        "pageThreeUnitPrice": 4,
    }

    df = google_auth.open_by_url(formURL)
    df = df[index_list_dict[df_name]].get_as_df()
    df = df[df.iloc[:,0].astype(str).str.len() > 0]

    if df_name == "breakages_df":
        df = df[df["选择操作"] == "录入破损物品"]
        df.drop(drop_list_dict[df_name], axis=1, inplace=True)
        df.columns = column_list_dict[df_name]
        df["日期(年年年年-月月-日日)"] = df["日期(年年年年-月月-日日)"].astype(str)
        df["日期(年年年年-月月-日日)"] = df["日期(年年年年-月月-日日)"].apply(lambda x: dt.datetime.strptime(x, "%m/%d/%Y").strftime("%Y-%m-%d"))
        df["日期(年年年年-月月-日日)"] = pd.to_datetime(df["日期(年年年年-月月-日日)"], format="%Y-%m-%d")

    elif df_name == "buyInStockDf":
        df = df[df["选择操作"]== "物品进货"]
        df.drop(drop_list_dict[df_name], axis=1, inplace=True)
        df.columns = column_list_dict[df_name]
        df["进货日期(年年年年-月月-日日)"] = df["进货日期(年年年年-月月-日日)"].astype(str)
        df["进货日期(年年年年-月月-日日)"] = df["进货日期(年年年年-月月-日日)"].apply(lambda x: dt.datetime.strptime(x, "%m/%d/%Y").strftime("%Y-%m-%d"))
        df["进货日期(年年年年-月月-日日)"] = pd.to_datetime(df["进货日期(年年年年-月月-日日)"], format="%Y-%m-%d")

    elif df_name == "lockInStockDf":
        pass

    elif df_name == "stockCountDf":
        df = df[df["选择操作"] == "盘点"]

        deleteTitleColumns = []
        for title in df.columns:
            if title not in ["Timestamp", "盘点日期"]:
                if "_盘点数量" not in title:
                    deleteTitleColumns += [title]
                else:
                    continue
            else:
                continue

        df.drop(deleteTitleColumns, axis=1, inplace=True)
        df["盘点日期"] = df["盘点日期"].astype(str)
        df["盘点日期"] = df["盘点日期"].apply(lambda x: dt.datetime.strptime(x, "%m/%d/%Y").strftime("%Y-%m-%d"))
        df["盘点日期"] = pd.to_datetime(df["盘点日期"], format="%Y-%m-%d")
        df.sort_values(by=["盘点日期"], ascending=True, ignore_index=True, inplace=True)

    elif df_name in ["pageOneUnitPrice", "pageTwoUnitPrice", "pageThreeUnitPrice"]:
        pass

    else:
        pass
    
    df.reset_index(inplace=True)
    df.drop("index", axis=1, inplace=True)

    return df

def is_leap_year(year):
    year = int(year)

    if (year % 400 == 0) and (year % 100 == 0):
        return True

    elif (year % 4 ==0) and (year % 100 != 0):
        return True

    else:
        return False

def month_last_day(year, month):
    year = int(year)
    month = int(month)

    parse_dict = {
        1 : 31,
        2 : 28,
        3 : 31,
        4 : 30,
        5 : 31,
        6 : 30,
        7 : 31,
        8 : 31,
        9 : 30,
        10 : 31,
        11 : 30,
        12 : 31,
    }

    if is_leap_year(year):
        parse_dict[2] = 29

    return int(parse_dict[month])

def on_internet():
    url = "https://www.google.com.sg/"
    timeout = 5
    try:
        request = requests.get(url, timeout=timeout)
        return True
    except (requests.ConnectionError, requests.Timeout) as exception:
        return False

def intRange(integer, lower, upper):
    if integer_check(integer):
        integer = int(integer)
        while integer < lower:
            print("输入的{}无效! 不能小于{}".format(integer, lower))
            integer = need_integer(input('请重新输入: '))
        else:
            while integer > upper:
                print("输入的{}无效! 不能大于{}".format(integer, upper))
                integer = need_integer(input('请重新输入: '))
            else:
                return integer
    else:
        while not integer_check(integer):
            print("输入的{}不是整数，请重新输入!".format(integer))
            integer = need_integer(input(': '))
        else:
            integer = int(integer)
            while integer < lower:
                print("输入的{}无效! 不能小于{}".format(integer, lower))
                integer = need_integer(input('请重新输入: '))
            else:
                while integer > upper:
                    print("输入的{}无效! 不能大于{}".format(integer, upper))
                    integer = need_integer(input('请重新输入: '))
                else:
                    return integer

def get_range_date(whole_month):
    if whole_month:
        inputYear = intRange(integer=input("输入年份: "), lower=2000, upper=2037)
        inputMonth = intRange(integer=input("输入月份: "), lower=1, upper=12)

        startDate = dt.datetime(year=int(inputYear), month=int(inputMonth), day=1)
        endDate = dt.datetime(year=int(inputYear), month=int(inputMonth), day=month_last_day(inputYear, inputMonth))

        startDate = pd.to_datetime(startDate)
        endDate = pd.to_datetime(endDate)

        return startDate, endDate

    else:
        startDate = None
        endDate = None

        while startDate == None or endDate == None:
            try:
                customStartYear = intRange(integer=input("请输入开始年份: "), lower=2000, upper=2037)
                customStartMonth = intRange(integer=input("请输入开始月份: "), lower=1, upper=12)
                customStartDay = intRange(integer=input("请输入开始日: "), lower=1, upper=31)

                startDate = dt.datetime(year=int(customStartYear), month=int(customStartMonth), day=int(customStartDay))

            except ValueError:
                print("开始日期错误, 请重新输入! ")
                startDate = None

            if startDate != None:
                try:
                    customEndYear = intRange(integer=input("请输入结束年份: "), lower=2000, upper=2037)
                    customEndMonth = intRange(integer=input("请输入结束月份: "), lower=1, upper=12)
                    customEndDay = intRange(integer=input("请输入结束日: "), lower=1, upper=31)

                    endDate = dt.datetime(year=int(customEndYear), month=int(customEndMonth), day=int(customEndDay))

                except ValueError:
                    print("结束日期错误, 请重新输入! ")
                    endDate = None
            else:
                endDate = None

            if endDate != None:
                if endDate < startDate:
                    print("结束日期不可以小于开始日期！")
                    endDate = None
        else:
            startDate = pd.to_datetime(startDate)
            endDate = pd.to_datetime(endDate)
            return startDate, endDate

def dfDateRangeFilter(df, df_column_name, print_string, startDate, endDate):
    customDateRangeFilter = pd.date_range(start=startDate, end=endDate)
    df[df_column_name] = pd.to_datetime(df[df_column_name])

    if df[df[df_column_name].isin(customDateRangeFilter)].empty:
        print("{}至{}无任何{}".format(startDate.strftime("%Y-%m-%d"), endDate.strftime("%Y-%m-%d"), print_string))
    else:
        print("{}至{}的{}".format(startDate.strftime("%Y-%m-%d"), endDate.strftime("%Y-%m-%d"), print_string))
        prtdf(df[df[df_column_name].isin(customDateRangeFilter)].sort_values(by=[df_column_name], ascending=False, ignore_index=True))

def df_remove_units(regex, original):
    reStr = re.search(regex, original)
    if reStr is None:
        return reStr
    else:
        return original.replace(reStr[0], '')

def doc_filename_check(fileName):
    regexYear = r'.*\年'
    regexMonth = r'年.*月'
    if re.search(regexYear, fileName) is None:
        raise TypeError
    else:
        if re.search(regexMonth, fileName) is None:
            raise TypeError
        else:
            return [re.search(regexYear, fileName)[0].replace("年", ""), re.search(regexMonth, fileName)[0].replace("月", "").replace("年", "")]

def stock_gen_pdf(customFont, pageOneDf, pageTwoDf, pageThreeDf, fileNameAfterCheck, outlet, stock_count_foldername, stock_count_pdf_foldername):

    outlet = str(outlet).strip().capitalize()

    with tqdm(total=100) as pbar:
        pbar.set_description("PDF生成中...")

        Document = borb_Document()
        Page = borb_Page(width=Decimal(595), height=Decimal(842))
        Document.add_page(Page)

        layout: borb_PageLayout = borb_SCL(Page)

        layout.add(borb_Paragraph("{}年{}月三人行({})楼面盘点记录表".format(fileNameAfterCheck[0], fileNameAfterCheck[1], outlet),
                                font=customFont,
                                horizontal_alignment=borb_align.CENTERED))

        table_1 = borb_flex_Table(number_of_rows=33, number_of_columns=len(pageOneDf.columns))

        for h in pageOneDf.columns:
            table_1.add(borb_TableCell(
                            borb_Paragraph(h, font=customFont, horizontal_alignment=borb_align.CENTERED)))

        space_avail = 32*len(pageOneDf.columns)
        appended_indexes = []
        for row in range(len(pageOneDf)):
            for column in range(len(pageOneDf.columns)):
                if space_avail > 0:
                    if str(pageOneDf.iloc[row, column]) == 'nan':
                        table_1.add(borb_Paragraph(' ', font=customFont, horizontal_alignment=borb_align.CENTERED))
                    else:
                        table_1.add(borb_Paragraph(str(pageOneDf.iloc[row,column]),
                                                font=customFont,
                                                horizontal_alignment=borb_align.CENTERED))

                    appended_indexes += [pageOneDf.iloc[row, 0]]
                    space_avail -= 1
                else:
                    continue

        max_index_appended = max(np.unique(appended_indexes))-1

        table_1.set_padding_on_all_cells(Decimal(1), Decimal(1), Decimal(1), Decimal(1))
        layout.add(table_1)

        pbar.update(20)

        table_2 = borb_flex_Table(number_of_rows=33, number_of_columns=len(pageOneDf.columns))

        for h in pageOneDf.columns:
            table_2.add(
                        borb_TableCell(
                        borb_Paragraph(h, font=customFont, horizontal_alignment=borb_align.CENTERED)))

        space_avail = 32*len(pageOneDf.columns)
        for row in range(max_index_appended+1, len(pageOneDf)):
            for col in range(len(pageOneDf.columns)):
                if space_avail > 0:
                    if str(pageOneDf.iloc[row, col]) == 'nan':
                        table_2.add(borb_Paragraph(" ", font=customFont, horizontal_alignment=borb_align.CENTERED))
                    else:
                        table_2.add(borb_Paragraph(str(pageOneDf.iloc[row, col]),
                                                font=customFont,
                                                horizontal_alignment=borb_align.CENTERED))
                    space_avail -= 1
                else:
                    continue

        for space in range(space_avail-1):
            table_2.add(borb_Paragraph(' ', font=customFont, horizontal_alignment=borb_align.CENTERED))

        table_2.set_padding_on_all_cells(Decimal(1), Decimal(1), Decimal(1), Decimal(1))
        layout.add(borb_Paragraph(' ', font=customFont))
        layout.add(table_2)
        layout.add(borb_Paragraph(' ', font=customFont))
        layout.add(borb_Paragraph(' ', font=customFont))

        pbar.update(20)

        layout.add(borb_Paragraph("{}年{}月三人行({})打包盒盘点记录表".format(fileNameAfterCheck[0], fileNameAfterCheck[1], outlet),
                                font=customFont,
                                horizontal_alignment=borb_align.CENTERED))

        table_3 = borb_flex_Table(number_of_rows=33, number_of_columns=len(pageTwoDf.columns))

        for h in pageTwoDf.columns:
            table_3.add(
                borb_TableCell(
                    borb_Paragraph(h, font=customFont, horizontal_alignment=borb_align.CENTERED),
                    )
                )

        space_avail = 32*len(pageTwoDf.columns)
        appended_indexes = []

        for row in range(len(pageTwoDf)):
            for col in range(len(pageTwoDf.columns)):
                if space_avail > 0:
                    if str(pageTwoDf.iloc[row, col]) == 'nan':
                        table_3.add(borb_Paragraph(' ', font=customFont))
                    else:
                        table_3.add(borb_Paragraph(str(pageTwoDf.iloc[row, col]), font=customFont, horizontal_alignment=borb_align.CENTERED))

                    appended_indexes += [pageTwoDf.iloc[row, 0]]
                    space_avail -= 1
                else:
                    continue

        max_index_appended = max(np.unique(appended_indexes))-1
        table_3.set_padding_on_all_cells(Decimal(1), Decimal(1), Decimal(1), Decimal(1))

        layout.add(table_3)
        layout.add(borb_Paragraph(' ', font=customFont))

        pbar.update(20)

        table_4 = borb_flex_Table(number_of_rows=33, number_of_columns=len(pageTwoDf.columns))

        for h in pageTwoDf.columns:
            table_4.add(
                borb_TableCell(
                    borb_Paragraph(h, font=customFont, horizontal_alignment=borb_align.CENTERED),
                    )
                )

        space_avail = 32*len(pageTwoDf.columns)
        for row in range(max_index_appended+1, len(pageTwoDf)):
            for col in range(len(pageTwoDf.columns)):
                if space_avail > 0:
                    if str(pageTwoDf.iloc[row, col]) == 'nan':
                        table_4.add(borb_Paragraph(' ', font=customFont))
                    else:
                        table_4.add(borb_Paragraph(str(pageTwoDf.iloc[row, col]), font=customFont, horizontal_alignment=borb_align.CENTERED))

                    space_avail -= 1
                else:
                    continue

        for space in range(space_avail-1):
            table_4.add(borb_Paragraph(' ', font=customFont, horizontal_alignment=borb_align.CENTERED))

        table_4.set_padding_on_all_cells(Decimal(1), Decimal(1), Decimal(1), Decimal(1))
        layout.add(table_4)
        layout.add(borb_Paragraph(' ', font=customFont))
        layout.add(borb_Paragraph(' ', font=customFont))

        pbar.update(20)

        layout.add(borb_Paragraph("{}年{}月三人行({})收银盘点记录表".format(fileNameAfterCheck[0], fileNameAfterCheck[1], outlet),
                                font=customFont,
                                horizontal_alignment=borb_align.CENTERED))

        table_5 = borb_flex_Table(number_of_rows=33, number_of_columns=len(pageThreeDf.columns))

        for h in pageThreeDf.columns:
            table_5.add(
                borb_TableCell(
                    borb_Paragraph(h, font=customFont, horizontal_alignment=borb_align.CENTERED),
                    )
                )

        space_avail = 32*len(pageThreeDf.columns)

        for row in range(len(pageThreeDf)):
            for col in range(len(pageThreeDf.columns)):
                if space_avail > 0:
                    if str(pageThreeDf.iloc[row, col]) == 'nan':
                        table_5.add(borb_Paragraph(' ', font=customFont))
                    else:
                        table_5.add(borb_Paragraph(str(pageThreeDf.iloc[row, col]), font=customFont, horizontal_alignment=borb_align.CENTERED))

                    space_avail -= 1
                else:
                    continue

        for space in range(space_avail-1):
            table_5.add(borb_Paragraph(' ', font=customFont, horizontal_alignment=borb_align.CENTERED))

        table_5.set_padding_on_all_cells(Decimal(1), Decimal(1), Decimal(1), Decimal(1))
        layout.add(table_5)

        pbar.set_description("PDF保存中...")
        pbar.update(15)

        pdf_filename = "{}年{}月{}盘点详情PDF.pdf".format(fileNameAfterCheck[0], fileNameAfterCheck[1], outlet.strip().capitalize())

        with open('{}/{}/{}/{}'.format(os.getcwd(), stock_count_foldername, stock_count_pdf_foldername, pdf_filename) , "wb") as pdf_file_handle:
                borb_PDF.dumps(pdf_file_handle, Document)

        pbar.set_description("PDF任务完成")
        pbar.update(5)

def inventory_main(google_auth, db_setting_url, constants_sheetname, serialized_rule_filename, backup_foldername, database_url, box_num, drink_num):
    google_auth = google_auth
    db_setting_url = db_setting_url
    constants_sheetname = constants_sheetname
    serialized_rule_filename = serialized_rule_filename
    backup_foldername = backup_foldername
    database_url = database_url
    box_num = box_num
    drink_num = drink_num

    #stock count Constants
    #lock in stock df for page one drop list
    LISDFFPOD_dict = {
        "Thomson" : ['23下','23壁橱', '66下', '68下','15下','28下','28壁橱','63壁橱','61壁橱','传菜口','制冰机上',88],
        "Bugis" : ["26下", "11下", "82壁橱", "83壁橱"],
    }

    songti_filename = "SongTi.ttf"
    stock_count_pdf_foldername = "PDF"

    res = pyfiglet.figlet_format("Inventory")
    print(res)
    time.sleep(0.15)

    on_net = on_internet()

    if not on_net:
        print()
        print()
        print("无网络连接。")
        print("库存盘点生成需要全程连接网络来完成。")

    else:
        fernet_key = get_key()

        if fernet_key == 0:
            print("安全密钥错误! ")
        else:
            storeOutlet = get_outlet()

            stock_count_foldername = "{}盘点文件".format(storeOutlet.strip().capitalize())
            stock_count_excel_foldername = "{}盘点详情EXCEL".format(storeOutlet.strip().capitalize())

            if not os.path.exists("{}/{}".format(os.getcwd(), stock_count_foldername)):
                os.makedirs("{}/{}".format(os.getcwd(), stock_count_foldername))
                print("检测出未创建盘点文件夹, 已自动创建盘点文件夹")

            if not os.path.exists("{}/{}/{}".format(os.getcwd(), stock_count_foldername, stock_count_excel_foldername)):
                os.makedirs("{}/{}/{}".format(os.getcwd(), stock_count_foldername, stock_count_excel_foldername))
                print("检测出未创建盘点详情EXCEL文件夹, 已自动创建盘点详情EXCEL文件夹")

            if not os.path.exists("{}/{}/{}".format(os.getcwd(), stock_count_foldername, stock_count_pdf_foldername)):
                os.makedirs("{}/{}/{}".format(os.getcwd(), stock_count_foldername, stock_count_pdf_foldername))
                print("检测未创建盘点PDF文件夹,已自动创建盘点PDF文件夹")

            with tqdm(total=100) as pbar1:
                pbar1.set_description("初始化...")

                k_dict = get_k_dictionary(google_auth, db_setting_url, constants_sheetname, serialized_rule_filename, storeOutlet, fernet_key, backup_foldername)
                formURL = fernet_decrypt(k_dict["inv_form_url"], fernet_key)

                pbar1.update(33)

                box_name_from_rule = []
                for i in range(box_num):
                    box_name = k_dict["box{}".format(i)]
                    box_name_from_rule += [box_name]

                drink_name_from_rule = []
                for j in range(drink_num):
                    drink_name = k_dict["drink{}".format(j)]
                    drink_name_from_rule += [drink_name]

                pbar1.update(33)

                try:
                    form_sheet = google_auth.open_by_url(formURL)
                    form_boxname_df = form_sheet[3].get_as_df()

                    form_boxname_dropIndex = form_boxname_df[form_boxname_df["物品"] == "BREAKLINE"].index[0]
                    form_boxname_df.drop(np.arange(form_boxname_dropIndex, len(form_boxname_df)), axis=0, inplace=True)

                    curr_box_name = form_boxname_df["物品"].to_list()

                    dr = pygsheets.datarange.DataRange(start="A2", end="A41", worksheet=form_sheet[3])

                    if len(curr_box_name) == len(box_name_from_rule):
                        if curr_box_name != box_name_from_rule:
                            update_matrix = []
                            for item in box_name_from_rule:
                                update_matrix += [[item]]

                            dr.update_values(values = update_matrix)
                            print("BOX NAME had been automatically updated with accordance to the Rule")
                            continue_stock_count = True
                        else:
                            continue_stock_count = True
                    else:
                        print("BOX NAME in form URL aren't the same length with the Rule")
                        continue_stock_count = False

                    form_drinkname_df = form_sheet[4].get_as_df()

                    form_drinkname_dropIndex = form_drinkname_df[form_drinkname_df["物品"] == "BREAKLINE"].index[0]
                    form_drinkname_df.drop(np.arange(form_drinkname_dropIndex, len(form_drinkname_df)), axis=0, inplace=True)

                    curr_drink_name = form_drinkname_df["物品"].to_list()

                    dr = pygsheets.datarange.DataRange(start="A2", end="A13", worksheet=form_sheet[4])

                    if len(curr_drink_name) == len(drink_name_from_rule):
                        if curr_drink_name != drink_name_from_rule:
                            update_matrix = []
                            for item in drink_name_from_rule:
                                update_matrix += [[item]]

                            dr.update_values(values = update_matrix)
                            print("DRINK NAME had been automatically updated with accordance to the Rule")
                        else:
                            pass
                    else:
                        print("DRINK NAME in form URL aren't the same length with the Rule")
                        continue_stock_count = False

                except Exception as e:
                    print()
                    print("读取或重置打包盒名称失败, 错误描述如下：")
                    print(e)
                    print()
                    try:
                        form_boxname_df = get_inv_df(formURL, "pageTwoUnitPrice")
                        form_boxname_dropIndex = form_boxname_df[form_boxname_df["物品"] == "BREAKLINE"].index[0]
                        form_boxname_df.drop(np.arange(form_boxname_dropIndex, len(form_boxname_df)), axis=0, inplace=True)

                        curr_box_name = form_boxname_df["物品"].to_list()

                        if len(curr_box_name) == len(box_name_from_rule):
                            if curr_box_name != box_name_from_rule:
                                print("打包盒名称与规则不同，盘点无法继续，请修改打包盒名称至与规则一模一样")
                                continue_stock_count = False
                            else:
                                continue_stock_count = True
                        else:
                            print("BOX NAME in form URL aren't the same length with the Rule")
                            continue_stock_count = False

                        form_drinkname_df = get_inv_df(formURL, "pageThreeUnitPrice")
                        form_drinkname_dropIndex = form_drinkname_df[form_drinkname_df["物品"] == "BREAKLINE"].index[0]
                        form_drinkname_df.drop(np.arange(form_drinkname_dropIndex, len(form_drinkname_df)), axis=0, inplace=True)

                        curr_drink_name = form_drinkname_df["物品"].to_list()

                        if len(curr_drink_name) == len(drink_name_from_rule):
                            if curr_drink_name != drink_name_from_rule:
                                print("酒水名称与规则不同，盘点无法继续，请修改打包盒名称至与规则一模一样")
                                continue_stock_count = False
                            else:
                                pass
                        else:
                            print("DRINK NAME in form URL aren't the same length with the Rule")
                            continue_stock_count = False

                    except Exception as e:
                        print()
                        print("通过仅读模式获取打包盒名称也失败了，错误描述如下：")
                        print(e)
                        continue_stock_count = False

                pbar1.set_description("初始化完成")
                pbar1.update(34)

            if continue_stock_count:
                userInputOne = 0
                while userInputOne != 3:
                    with tqdm(total=100) as pbar2:
                        pbar2.set_description("刷新文件...")

                        breakages_df = get_inv_df(formURL, "breakages_df")
                        pbar2.update(25)

                        buyInStockDf = get_inv_df(formURL, "buyInStockDf")
                        pbar2.update(25)

                        lockInStockDf = get_inv_df(formURL, "lockInStockDf")
                        pbar2.update(25)

                        stockCountDf = get_inv_df(formURL, "stockCountDf")
                        pbar2.update(24)

                        pbar2.set_description("刷新完成")
                        pbar2.update(1)

                    print("盘点主菜单")
                    startAction = option_num(["查看文档", "盘点类操作", "生成PDF", "退出盘点"])
                    time.sleep(0.25)
                    userInputOne = option_limit(startAction, input("在这里输入>>>: "))

                    if userInputOne == 0:
                        userInputTwo = 0
                        while userInputTwo != 3:
                            docActions = option_num(["查看物品破损记录表", "查看物品进货表", "查看封存区库存", "返回上一菜单"])
                            time.sleep(0.25)
                            userInputTwo = option_limit(docActions, input("在这里输入>>>: "))

                            if userInputTwo == 0:
                                userInputThree = 0
                                while userInputThree != 3:
                                    breakagesDocActions = option_num(['整月查看','自定义日期范围', '查看全部破损记录', '返回上一菜单'])
                                    time.sleep(0.25)
                                    userInputThree = option_limit(breakagesDocActions, input("在这里输入>>>: "))

                                    if userInputThree == 0:

                                        startDate, endDate = get_range_date(whole_month=True)

                                        dfDateRangeFilter(df=breakages_df, df_column_name="日期(年年年年-月月-日日)", print_string="破损记录", startDate=startDate, endDate=endDate)

                                    elif userInputThree == 1:

                                        startDate, endDate = get_range_date(whole_month=False)

                                        dfDateRangeFilter(df=breakages_df, df_column_name="日期(年年年年-月月-日日)", print_string="破损记录", startDate=startDate, endDate=endDate)

                                    elif userInputThree == 2:
                                        print("显示全部破损记录")
                                        breakages_df["日期(年年年年-月月-日日)"] = pd.to_datetime(breakages_df["日期(年年年年-月月-日日)"])
                                        prtdf(breakages_df.sort_values(by=["日期(年年年年-月月-日日)"], ascending=False, ignore_index=True))

                            elif userInputTwo == 1:
                                userInputFour = 0
                                while userInputFour != 3:
                                    buyInStockActions = option_num(['整月查看','自定义日期范围', '查看全部入库记录', '返回上一菜单'])
                                    time.sleep(0.25)
                                    userInputFour = option_limit(buyInStockActions, input("在这里输入>>>: "))

                                    if userInputFour == 0:
                                        startDate, endDate = get_range_date(whole_month=True)

                                        dfDateRangeFilter(df=buyInStockDf, df_column_name="进货日期(年年年年-月月-日日)", print_string="进货记录", startDate=startDate, endDate=endDate)

                                    elif userInputFour == 1:
                                        startDate, endDate = get_range_date(whole_month=False)

                                        dfDateRangeFilter(df=buyInStockDf, df_column_name="进货日期(年年年年-月月-日日)", print_string="进货记录", startDate=startDate, endDate=endDate)

                                    elif userInputFour == 2:
                                        print("显示全部进货记录")
                                        buyInStockDf["进货日期(年年年年-月月-日日)"] = pd.to_datetime(buyInStockDf["进货日期(年年年年-月月-日日)"])
                                        prtdf(buyInStockDf.sort_values(by=["进货日期(年年年年-月月-日日)"], ascending=False, ignore_index=True))

                            elif userInputTwo == 2:
                                print("封存区库存")
                                prtdf(lockInStockDf)

                    elif userInputOne == 1:
                        userInputFive = 0
                        while userInputFive != 2:
                            stockRelatedAction = option_num(["开始盘点", "查看往月盘点详情", "返回上一菜单"])
                            time.sleep(0.25)
                            userInputFive = option_limit(stockRelatedAction, input("在这里输入>>>: "))

                            if userInputFive == 0:

                                stockStartDate, stockEndDate = get_range_date(whole_month=True)
                                dateRangeFilter = pd.date_range(start=stockStartDate, end=stockEndDate)

                                stockCountDfFiltered = stockCountDf.copy()
                                stockCountDfFiltered["盘点日期"] = pd.to_datetime(stockCountDfFiltered["盘点日期"])
                                stockCountDfFiltered = stockCountDfFiltered[stockCountDfFiltered["盘点日期"].isin(dateRangeFilter)]
                                stockCountDfFiltered.sort_values(by=["盘点日期"], ascending=True, ignore_index=True, inplace=True)

                                if stockCountDfFiltered.empty:
                                    print("你还没有录入{}年{}月的现有数量盘点，盘点无法继续".format(stockStartDate.year, stockStartDate.month))

                                else:
                                    stockCountDfFiltered = stockCountDfFiltered.iloc[-1, :]
                                    if stockCountDfFiltered.empty:
                                        print("Critical Error 关键错误")
                                    else:
                                        database_url = fernet_decrypt(database_url, fernet_key).strip()
                                        box_drink_in_out_url = fernet_decrypt(k_dict["box_drink_in_out_url"], fernet_key).strip()

                                        drink_db_sheetname = k_dict["drink_database_sheetname"]
                                        takeaway_box_sheetname = k_dict["takeaway_box_database_sheetname"]

                                        tabox_stock_sheetname = k_dict["takeaway_box_sheetname"]
                                        drink_stock_sheetname = k_dict["drink_stock_sheetname"]

                                        local_db_filename = k_dict["local_database_filename"]

                                        try:
                                            online_db_sheet = google_auth.open_by_url(database_url)
                                            drink_db_sheetname_index = online_db_sheet.worksheet(property="title", value=drink_db_sheetname).index
                                            takeaway_box_sheetname_index = online_db_sheet.worksheet(property="title", value=takeaway_box_sheetname).index

                                            drink_db = online_db_sheet[drink_db_sheetname_index].get_as_df()
                                            tabox_db = online_db_sheet[takeaway_box_sheetname_index].get_as_df()

                                        except Exception as e:
                                            print()
                                            print("网络获取打包盒和酒水数据库失败，错误描述如下：")
                                            print(e)
                                            print()
                                            try:
                                                drink_db = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_db_filename), sheet_name=drink_db_sheetname)
                                                tabox_db = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_db_filename), sheet_name=takeaway_box_sheetname)

                                            except Exception as e:
                                                print()
                                                print("无法从本地获打包盒和酒水数据库，错误描述如下：")
                                                print(e)
                                                print()
                                                drink_db = None
                                                tabox_db = None

                                        try:
                                            box_drink_sheet = google_auth.open_by_url(box_drink_in_out_url)

                                            tabox_stock_sheetname_index = box_drink_sheet.worksheet(property="title", value=tabox_stock_sheetname).index
                                            drink_sheetname_index = box_drink_sheet.worksheet(property="title", value=drink_stock_sheetname).index

                                            tabox_stock_df = box_drink_sheet[tabox_stock_sheetname_index].get_as_df()
                                            drink_stock_df = box_drink_sheet[drink_sheetname_index].get_as_df()

                                        except Exception as e:
                                            print()
                                            print("通过机器人从网络获取打包盒和酒水出入库失败，错误描述如下： ")
                                            print(e)
                                            print()
                                            tabox_stock_df = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_db_filename), sheet_name=tabox_stock_sheetname)
                                            drink_stock_df = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_db_filename), sheet_name=drink_stock_sheetname)
                                            print()
                                            print("已采用本地备份打包盒和酒水出入库信息")

                                        if isinstance(drink_db, pd.DataFrame) and isinstance(tabox_db, pd.DataFrame):

                                            previousDayFromStockDate = dt.datetime(year=stockStartDate.year, month=stockStartDate.month, day=stockStartDate.day) - dt.timedelta(days=1)
                                            previousStockCountFilename = "{}年{}月{}盘点详情.xlsx".format(previousDayFromStockDate.year, previousDayFromStockDate.month, storeOutlet.strip().capitalize())

                                            if not os.path.exists("{}/{}/{}/{}".format(os.getcwd(), stock_count_foldername, stock_count_excel_foldername, previousStockCountFilename)):
                                                sheet1_columns = ["编号", "物品(单位)", "单价", "上月存货", "进货数量", "破损数量", "损耗数量", "现有数量", "备注"]
                                                sheet2_columns = ["编号", "物品", "单位", "单价", "上月存货", "进货数量", "本月使用量", "现有数量", "备注"]
                                                sheet3_columns = ["编号", "物品", "单位", "单价", "上月存货", "进货数量", "本月使用量", "现有数量", "备注"]

                                                pageOneUnitPrice = get_inv_df(formURL, "pageOneUnitPrice")
                                                pageTwoUnitPrice = get_inv_df(formURL, "pageTwoUnitPrice")
                                                pageThreeUnitPrice = get_inv_df(formURL, "pageThreeUnitPrice")

                                                pageTwoUnitPrice["ACTIVE STATUS"] = pageTwoUnitPrice["ACTIVE STATUS"].astype(int)
                                                pageThreeUnitPrice["ACTIVE STATUS"] =  pageThreeUnitPrice["ACTIVE STATUS"].astype(int)

                                                sheet1_items = pageOneUnitPrice["物品(单位)"].values

                                                sheet1_price = pageOneUnitPrice["单价"].values

                                                dummy_sheet1 = pd.DataFrame({
                                                        sheet1_columns[0] : np.arange(1, len(sheet1_items)+1),
                                                        sheet1_columns[1] : sheet1_items,
                                                        sheet1_columns[2] : sheet1_price,
                                                        sheet1_columns[3] : np.zeros(len(sheet1_items)),
                                                        sheet1_columns[4] : np.zeros(len(sheet1_items)),
                                                        sheet1_columns[5] : np.zeros(len(sheet1_items)),
                                                        sheet1_columns[6] : np.zeros(len(sheet1_items)),
                                                        sheet1_columns[7] : np.zeros(len(sheet1_items)),
                                                        sheet1_columns[8] : np.repeat(None, len(sheet1_items)),
                                                    })


                                                sheet2_items = []

                                                for index in range(len(tabox_stock_df)):
                                                    if pageTwoUnitPrice.iloc[index, 3] == 1:
                                                        sheet2_items += [str(tabox_stock_df.iloc[index, 10])]
                                                    else:
                                                        continue

                                                for index in range(len(pageTwoUnitPrice)-len(tabox_stock_df)):
                                                    i = index + len(tabox_stock_df)
                                                    if pageTwoUnitPrice.iloc[i, 3] == 1:
                                                        sheet2_items += [pageTwoUnitPrice.iloc[i, 0]]
                                                    else:
                                                        continue

                                                sheet2_unit = pageTwoUnitPrice[pageTwoUnitPrice["ACTIVE STATUS"] == 1]["单位"].values

                                                sheet2_price = pageTwoUnitPrice[pageTwoUnitPrice["ACTIVE STATUS"] == 1]["单价"].values

                                                dummy_sheet2 = pd.DataFrame({
                                                            sheet2_columns[0] : np.arange(1, len(sheet2_items)+1),
                                                            sheet2_columns[1] : sheet2_items,
                                                            sheet2_columns[2] : sheet2_unit,
                                                            sheet2_columns[3] : sheet2_price,
                                                            sheet2_columns[4] : np.zeros(len(sheet2_items)),
                                                            sheet2_columns[5] : np.zeros(len(sheet2_items)),
                                                            sheet2_columns[6] : np.zeros(len(sheet2_items)),
                                                            sheet2_columns[7] : np.zeros(len(sheet2_items)),
                                                            sheet2_columns[8] : np.repeat(None, len(sheet2_items)),
                                                    })

                                                sheet3_items = []

                                                for index in range(len(drink_stock_df)):
                                                    if pageThreeUnitPrice.iloc[index, 3] == 1:
                                                        sheet3_items += [str(drink_stock_df.iloc[index, 7])]
                                                    else:
                                                        continue

                                                for index in range(len(pageThreeUnitPrice) - len(drink_stock_df)):
                                                    i  = index + len(drink_stock_df)

                                                    if pageThreeUnitPrice.iloc[i, 3] == 1:
                                                        sheet3_items += [pageThreeUnitPrice.iloc[i, 0]]
                                                    else:
                                                        continue

                                                sheet3_unit = pageThreeUnitPrice[pageThreeUnitPrice["ACTIVE STATUS"] == 1]["单位"].values

                                                sheet3_price = pageThreeUnitPrice[pageThreeUnitPrice["ACTIVE STATUS"] == 1]["单价"].values

                                                dummy_sheet3 = pd.DataFrame({
                                                            sheet3_columns[0] : np.arange(1, len(sheet3_items)+1),
                                                            sheet3_columns[1] : sheet3_items,
                                                            sheet3_columns[2] : sheet3_unit,
                                                            sheet3_columns[3] : sheet3_price,
                                                            sheet3_columns[4] : np.zeros(len(sheet3_items)),
                                                            sheet3_columns[5] : np.zeros(len(sheet3_items)),
                                                            sheet3_columns[6] : np.zeros(len(sheet3_items)),
                                                            sheet3_columns[7] : np.zeros(len(sheet3_items)),
                                                            sheet3_columns[8] : np.repeat(None, len(sheet3_items)),
                                                    })


                                                with pd.ExcelWriter('{}/{}/{}/{}'.format(os.getcwd(), stock_count_foldername, stock_count_excel_foldername, previousStockCountFilename)) as writer:
                                                    dummy_sheet1.to_excel(writer, sheet_name="Sheet1", header=True, index=False)
                                                    dummy_sheet2.to_excel(writer, sheet_name="Sheet2", header=True, index=False)
                                                    dummy_sheet3.to_excel(writer, sheet_name="Sheet3", header=True, index=False)

                                                print("未发现上月盘点Excel, 已创建'{}'伪文件以继续本次的盘点".format(previousStockCountFilename))

                                            else:
                                                pass


                                            with tqdm(total=100) as pbar3:
                                                pbar3.set_description("盘点计算中...")
                                                stockCountDfIndexChangeOne = [i.replace("_盘点数量","") for i in stockCountDfFiltered.index]
                                                stockCountDfFilteredForPageOne = stockCountDfFiltered.copy()
                                                stockCountDfFilteredForPageOne.index = stockCountDfIndexChangeOne

                                                regex = r'\(.*\)'
                                                stockCountDfFilteredForPageOneIndexes = stockCountDfFilteredForPageOne.copy().index.values
                                                stockCountDfFilteredRemoveUnitsArray = np.repeat(None, len(stockCountDfFilteredForPageOneIndexes))
                                                for name in range(len(stockCountDfFilteredForPageOneIndexes)):
                                                    reStr = re.search(regex, stockCountDfFilteredForPageOneIndexes[name])
                                                    if reStr is None:
                                                        stockCountDfFilteredRemoveUnitsArray[name] = stockCountDfFilteredForPageOneIndexes[name]
                                                    else:
                                                        stockCountDfFilteredRemoveUnitsArray[name] = stockCountDfFilteredForPageOneIndexes[name].replace(reStr[0], '')

                                                stockCountDfFilteredRemoveUnits = stockCountDfFilteredForPageOne.copy()
                                                stockCountDfFilteredRemoveUnits.index = stockCountDfFilteredRemoveUnitsArray

                                                buyInStockDfFiltered = buyInStockDf.copy()
                                                buyInStockDfFiltered = buyInStockDfFiltered[buyInStockDfFiltered['进货日期(年年年年-月月-日日)'].isin(dateRangeFilter)]

                                                if buyInStockDfFiltered.empty:
                                                    buyInStockDfFilteredGrouped = pd.DataFrame(columns=['物品(单位)', '进货数量'])

                                                else:
                                                    buyInStockDfFiltered["进货数量"] = buyInStockDfFiltered["进货数量"].astype(int)
                                                    buyInStockDfFilteredGrouped = buyInStockDfFiltered.groupby("进货物品(单位)")[['进货数量']].sum()
                                                    buyInStockDfFilteredGrouped.reset_index(inplace=True)
                                                    buyInStockDfFilteredGrouped.columns = ['物品(单位)', '进货数量']


                                                buyInStockDfFilteredGroupedRemoveUnits = buyInStockDfFilteredGrouped.copy()
                                                buyInStockDfFilteredGroupedRemoveUnits['物品(单位)'] = buyInStockDfFilteredGroupedRemoveUnits['物品(单位)'].apply(lambda x: df_remove_units(regex, x))

                                                breakages_dfFiltered = breakages_df.copy()
                                                breakages_dfFiltered = breakages_dfFiltered[breakages_dfFiltered["日期(年年年年-月月-日日)"].isin(dateRangeFilter)]

                                                if breakages_dfFiltered.empty:
                                                    breakages_dfFilteredGrouped = pd.DataFrame(columns=['物品(单位)', "破损数量"])

                                                else:
                                                    breakages_dfFiltered["破损数量"] = breakages_dfFiltered["破损数量"].astype(int)
                                                    breakages_dfFilteredGrouped = breakages_dfFiltered.groupby("破损物品(单位)")[['破损数量']].sum()
                                                    breakages_dfFilteredGrouped.reset_index(inplace=True)
                                                    breakages_dfFilteredGrouped.columns = ['物品(单位)', "破损数量"]

                                                pageOneUnitPrice = get_inv_df(formURL, "pageOneUnitPrice")

                                                pageOnePriceRawList = pageOneUnitPrice[pageOneUnitPrice.columns[1]].values
                                                pageOnePriceList=[' ' if x is np.nan else x for x in pageOnePriceRawList]

                                                previousStockCount = pd.read_excel("{}/{}/{}/{}".format(os.getcwd(), stock_count_foldername, stock_count_excel_foldername, previousStockCountFilename), sheet_name="Sheet1")
                                                previousStockCount.drop(['编号','单价', '上月存货', '进货数量', '破损数量', '损耗数量','备注'], axis=1, inplace=True)

                                                pageOneDf = pd.DataFrame({"编号" : np.arange(1, len(pageOneUnitPrice)+1),
                                                                    "物品(单位)": pageOneUnitPrice[pageOneUnitPrice.columns[0]].values,
                                                                    })

                                                previousStockArray = np.repeat(None, len(pageOneDf['物品(单位)'].values))
                                                for itemName in range(len(pageOneDf['物品(单位)'].values)):
                                                    try:
                                                        previousStockArray[itemName] = int(previousStockCount[previousStockCount['物品(单位)'] == pageOneDf['物品(单位)'].values[itemName]]["现有数量"].values[0])
                                                    except IndexError:
                                                        previousStockArray[itemName] = -404
                                                    except ValueError:
                                                        previousStockArray[itemName] = -406

                                                pageOneDf['上月存货'] = previousStockArray
                                                pageOneDf['上月存货'] = pageOneDf['上月存货'].astype(int)

                                                buyInStockDfFilteredGroupedForPageOne = buyInStockDfFilteredGrouped.copy()
                                                buyInStockDfFilteredGroupedForPageOneIndexDrop = []

                                                for item in range(len(buyInStockDfFilteredGroupedForPageOne)):
                                                    if buyInStockDfFilteredGroupedForPageOne['物品(单位)'].values[item] not in pageOneDf['物品(单位)'].values:
                                                        buyInStockDfFilteredGroupedForPageOneIndexDrop += [item]
                                                buyInStockDfFilteredGroupedForPageOne.drop(buyInStockDfFilteredGroupedForPageOneIndexDrop, axis=0, inplace=True)
                                                buyInStockDfFilteredGroupedForPageOne.reset_index(inplace=True)
                                                buyInStockDfFilteredGroupedForPageOne.drop("index", axis=1, inplace=True)

                                                pageOneDf = pd.merge(right=buyInStockDfFilteredGroupedForPageOne, left=pageOneDf, how="outer")

                                                breakages_dfFilteredGroupedForPageOne = breakages_dfFilteredGrouped.copy()
                                                breakages_dfFilteredGroupedForPageOneIndexDrop = []

                                                for item in range(len(breakages_dfFilteredGroupedForPageOne)):
                                                    if breakages_dfFilteredGroupedForPageOne['物品(单位)'].values[item] not in pageOneDf['物品(单位)'].values:
                                                        breakages_dfFilteredGroupedForPageOneIndexDrop += [item]

                                                breakages_dfFilteredGroupedForPageOne.drop(breakages_dfFilteredGroupedForPageOneIndexDrop, axis=0, inplace=True)
                                                breakages_dfFilteredGroupedForPageOne.reset_index(inplace=True)
                                                breakages_dfFilteredGroupedForPageOne.drop("index", axis=1, inplace=True)

                                                pageOneDf = pd.merge(right=breakages_dfFilteredGroupedForPageOne, left=pageOneDf, how="outer")

                                                breakagesColumnsTurnZero = [0 if str(x) == 'nan' else x for x in pageOneDf['破损数量'].values]
                                                pageOneDf['破损数量'] = breakagesColumnsTurnZero
                                                buyInStockColumnsTurnZero = [0 if str(x) == 'nan' else x for x in pageOneDf['进货数量'].values]
                                                pageOneDf['进货数量'] = buyInStockColumnsTurnZero

                                                lockInStockDfForPageOne = lockInStockDf.copy()

                                                lockInStockDfForPageOne.drop(LISDFFPOD_dict["{}".format(storeOutlet.strip().capitalize())], axis=1, inplace=True)

                                                lockInStockDfForPageOne.columns = ['物品(单位)', '总数']

                                                pageOneCurrentArray = np.repeat(None, len(pageOneDf))

                                                for num in range(len(pageOneDf)):
                                                    pageOneCurrentArray[num] = stockCountDfFilteredForPageOne[pageOneDf['物品(单位)'].values[num]]

                                                addLockInStockArray = np.repeat(None, len(pageOneDf))

                                                for num in range(len(pageOneDf)):
                                                    addLockInStockArray[num] = lockInStockDfForPageOne[lockInStockDfForPageOne['物品(单位)'] == pageOneDf['物品(单位)'].values[num]]['总数'].sum()

                                                PageOneDfCurrentwithLockInStock = np.add(pageOneCurrentArray.astype(int), addLockInStockArray.astype(int))

                                                pageOneDf['现有数量'] = PageOneDfCurrentwithLockInStock

                                                pageOneDf['破损数量'] = pageOneDf['破损数量'].astype(int)
                                                pageOneDf['进货数量'] = pageOneDf['进货数量'].astype(int)
                                                pageOneDf['现有数量'] = pageOneDf['现有数量'].astype(int)

                                                pageOneDf['损耗数量'] = pageOneDf['上月存货'] - pageOneDf['现有数量'] + pageOneDf['进货数量'] - pageOneDf['破损数量']

                                                pageOnePriceRestructArray = np.repeat(None, len(pageOnePriceList))

                                                for price in range(len(pageOnePriceList)):
                                                    if pageOnePriceList[price] == ' ':
                                                        pageOnePriceRestructArray[price] = ' '
                                                    else:
                                                        pageOnePriceRestructArray[price] = '${}'.format(format(float(pageOnePriceList[price]), '.2f'))


                                                pageOneDf['单价'] = pageOnePriceRestructArray
                                                pageOneDf = pageOneDf[['编号', '物品(单位)', '单价', '上月存货', '进货数量', '破损数量', '损耗数量','现有数量']]

                                                pbar3.update(33)

                                                pageTwoUnitPrice = get_inv_df(formURL, "pageTwoUnitPrice")

                                                pageTwoUnitPriceAuto = pageTwoUnitPrice.copy()
                                                pageTwoUnitPriceAutoDropIndex = pageTwoUnitPrice[pageTwoUnitPrice['物品'] == 'BREAKLINE'].index[0]
                                                pageTwoUnitPriceAuto.drop(np.arange(pageTwoUnitPriceAutoDropIndex, len(pageTwoUnitPrice)), axis=0, inplace=True)

                                                tabox_db["DATE"] = pd.to_datetime(tabox_db["DATE"])
                                                tabox_db = tabox_db[tabox_db["DATE"].isin(dateRangeFilter)]

                                                TBstockInColumnTitleForDrop = []
                                                for column in tabox_db.columns:
                                                    if column not in ["DATE"]:
                                                        if "入库" not in column:
                                                            TBstockInColumnTitleForDrop += [column]
                                                        else:
                                                            continue
                                                    else:
                                                        continue

                                                TBFileStockIn = tabox_db.copy()
                                                TBFileStockIn.drop(TBstockInColumnTitleForDrop, axis=1, inplace=True)

                                                if TBFileStockIn.empty:
                                                    TBstockInSums = np.repeat(-404, len(pageTwoUnitPriceAuto['物品'].values))
                                                    print("自动化打包盒在数据库里没有当月的入库信息，将以数值-404呈现")
                                                else:
                                                    TBstockInSums = np.repeat(None, len(pageTwoUnitPriceAuto['物品'].values))

                                                    for item in range(len(pageTwoUnitPriceAuto['物品'].values)):
                                                        TBstockInSums[item]  = np.floor(TBFileStockIn[pageTwoUnitPriceAuto['物品'].values[item]+'入库'].sum())

                                                TBstockInAuto = pd.DataFrame({'物品': pageTwoUnitPriceAuto['物品'].values,
                                                                            "进货数量": TBstockInSums.astype(int)})

                                                pageTwoUnitPriceAuto = pd.merge(right=TBstockInAuto, left = pageTwoUnitPriceAuto, how='outer')


                                                TBOutColumnTitleForDrop = []
                                                for column in tabox_db.columns:
                                                    if column not in ['DATE']:
                                                        if "出库" not in column:
                                                            TBOutColumnTitleForDrop += [column]
                                                        else:
                                                            continue
                                                    else:
                                                        continue

                                                TBFileOut = tabox_db.copy()
                                                TBFileOut.drop(TBOutColumnTitleForDrop, axis=1, inplace=True)

                                                if TBFileOut.empty:
                                                    TBOutSums = np.repeat(-404, len(pageTwoUnitPriceAuto['物品'].values))
                                                    print("自动化打包盒在数据库里没有当月的出库信息，将以数值-404呈现")
                                                else:
                                                    TBOutSums = np.repeat(None, len(pageTwoUnitPriceAuto['物品'].values))

                                                    for item in range(len(pageTwoUnitPriceAuto['物品'].values)):
                                                        TBOutSums[item]  = np.ceil(TBFileOut[pageTwoUnitPriceAuto['物品'].values[item]+'出库'].sum())

                                                TBOutAuto = pd.DataFrame({
                                                                        "编号": np.arange(1, len(pageTwoUnitPriceAuto['物品'].values)+1),
                                                                        "物品": pageTwoUnitPriceAuto['物品'].values,
                                                                        "本月使用量": TBOutSums.astype(int)
                                                                        })

                                                pageTwoUnitPriceAuto = pd.merge(right=TBOutAuto, left = pageTwoUnitPriceAuto, how='outer')


                                                TBCurrentColumnTitleForDrop = []
                                                for column in tabox_db.columns:
                                                    if column not in ['DATE']:
                                                        if "现有数量" not in column:
                                                            TBCurrentColumnTitleForDrop += [column]
                                                        else:
                                                            continue
                                                    else:
                                                        continue

                                                TBFileCurrent = tabox_db.copy()
                                                TBFileCurrent.drop(TBCurrentColumnTitleForDrop, axis=1, inplace=True)

                                                if TBFileCurrent.empty:
                                                    TBCurrentSums = np.repeat(-404, len(pageTwoUnitPriceAuto['物品'].values))
                                                    print("自动化打包盒在数据库里没有当月的完整信息，将以数值-404呈现")
                                                else:
                                                    TBFileCurrent["DATE"] = pd.to_datetime(TBFileCurrent["DATE"])
                                                    maxDateForTBFileCurrent = TBFileCurrent["DATE"].max()
                                                    TBFileCurrent = TBFileCurrent[TBFileCurrent['DATE'] == maxDateForTBFileCurrent]

                                                    TBCurrentSums = np.repeat(None, len(pageTwoUnitPriceAuto['物品'].values))
                                                    for item in range(len(pageTwoUnitPriceAuto['物品'].values)):
                                                        TBCurrentSums[item] = np.floor(TBFileCurrent[pageTwoUnitPriceAuto['物品'].values[item]+'现有数量'].sum())

                                                TBCurrentAuto = pd.DataFrame({"物品": pageTwoUnitPriceAuto['物品'].values,
                                                                            "现有数量": TBCurrentSums.astype(int)})

                                                pageTwoUnitPriceAuto = pd.merge(right=TBCurrentAuto, left = pageTwoUnitPriceAuto, how='outer')

                                                if len(pageTwoUnitPriceAuto) == len(tabox_stock_df):
                                                    nameReplace = np.repeat(None, len(pageTwoUnitPriceAuto))
                                                    for index in range(len(tabox_stock_df)):
                                                        nameReplace[index] = tabox_stock_df.iloc[index, 10]

                                                    pageTwoUnitPriceAuto["物品"] = nameReplace

                                                else:
                                                    print("Critical Error, Takeaway box names have not replaced for stock count")

                                                pageTwoUnitPriceAuto["ACTIVE STATUS"] = pageTwoUnitPriceAuto["ACTIVE STATUS"].astype(int)
                                                pageTwoUnitPriceAuto = pageTwoUnitPriceAuto[pageTwoUnitPriceAuto["ACTIVE STATUS"] == 1]

                                                previousStockCountPageTwo = pd.read_excel('{}/{}/{}/{}'.format(os.getcwd(), stock_count_foldername, stock_count_excel_foldername, previousStockCountFilename),sheet_name="Sheet2")
                                                previousStockCountPageTwo.drop(['编号','单位', '单价', '上月存货', '进货数量', '本月使用量','备注'], axis=1, inplace=True)

                                                pageTwoUnitPriceAutoPreviousStockArray = np.repeat(None, len(pageTwoUnitPriceAuto))
                                                for name in range(len(pageTwoUnitPriceAuto)):
                                                    try:
                                                        pageTwoUnitPriceAutoPreviousStockArray[name] = int(previousStockCountPageTwo[previousStockCountPageTwo['物品'] == pageTwoUnitPriceAuto['物品'].values[name]]['现有数量'].values[0])
                                                    except IndexError:
                                                        pageTwoUnitPriceAutoPreviousStockArray[name] = -404
                                                    except ValueError:
                                                        pageTwoUnitPriceAutoPreviousStockArray[name] = -406

                                                pageTwoUnitPriceAuto['上月存货'] = pageTwoUnitPriceAutoPreviousStockArray
                                                pageTwoUnitPriceAuto = pageTwoUnitPriceAuto[['编号','物品', '单位', '单价', "上月存货",'进货数量','本月使用量', '现有数量']]

                                                pageTwoUnitPriceNOTAuto = pageTwoUnitPrice.copy()
                                                pageTwoUnitPriceNOTAuto.drop(np.arange(0, pageTwoUnitPriceAutoDropIndex+1), axis=0, inplace=True)

                                                buyInStockDfFilteredGroupedRemoveUnitsForPageTwoNOTAuto = buyInStockDfFilteredGroupedRemoveUnits.copy()
                                                buyInStockDfFilteredGroupedRemoveUnitsForPageTwoNOTAutoIndexDrop = []

                                                for name in range(len(buyInStockDfFilteredGroupedRemoveUnitsForPageTwoNOTAuto)):
                                                    if buyInStockDfFilteredGroupedRemoveUnitsForPageTwoNOTAuto['物品(单位)'].values[name] not in pageTwoUnitPriceNOTAuto['物品'].values:
                                                        buyInStockDfFilteredGroupedRemoveUnitsForPageTwoNOTAutoIndexDrop += [name]
                                                    else:
                                                        continue

                                                buyInStockDfFilteredGroupedRemoveUnitsForPageTwoNOTAuto.drop(buyInStockDfFilteredGroupedRemoveUnitsForPageTwoNOTAutoIndexDrop, axis=0, inplace=True)
                                                buyInStockDfFilteredGroupedRemoveUnitsForPageTwoNOTAuto.reset_index(inplace=True)
                                                buyInStockDfFilteredGroupedRemoveUnitsForPageTwoNOTAuto.drop("index", axis=1, inplace=True)
                                                buyInStockDfFilteredGroupedRemoveUnitsForPageTwoNOTAuto.columns = ["物品", "进货数量"]

                                                pageTwoUnitPriceNOTAuto = pd.merge(right=buyInStockDfFilteredGroupedRemoveUnitsForPageTwoNOTAuto, left=pageTwoUnitPriceNOTAuto, how='outer')

                                                buyInStockDfFilteredGroupedRemoveUnitsForPageTwoNOTAutoTurnZero = [0 if str(x) == 'nan' else x for x in pageTwoUnitPriceNOTAuto['进货数量'].values]
                                                pageTwoUnitPriceNOTAuto['进货数量'] = buyInStockDfFilteredGroupedRemoveUnitsForPageTwoNOTAutoTurnZero
                                                pageTwoUnitPriceNOTAuto['进货数量'] = pageTwoUnitPriceNOTAuto['进货数量'].astype(int)

                                                pageTwoUnitPriceNOTAutoCurrentArray = np.repeat(None, len(pageTwoUnitPriceNOTAuto))

                                                for name in range(len(pageTwoUnitPriceNOTAuto)):
                                                    pageTwoUnitPriceNOTAutoCurrentArray[name] = stockCountDfFilteredRemoveUnits[pageTwoUnitPriceNOTAuto['物品'].values[name]]

                                                pageTwoUnitPriceNOTAuto['现有数量'] = pageTwoUnitPriceNOTAutoCurrentArray
                                                pageTwoUnitPriceNOTAuto['现有数量'] = pageTwoUnitPriceNOTAuto['现有数量'].apply(lambda x: np.floor(float(x)))
                                                pageTwoUnitPriceNOTAuto['现有数量'] = pageTwoUnitPriceNOTAuto['现有数量'].astype(int)

                                                pageTwoUnitPriceNOTAutoPreviousStockArray = np.repeat(None, len(pageTwoUnitPriceNOTAuto))

                                                for name in range(len(pageTwoUnitPriceNOTAuto)):
                                                    try:
                                                        pageTwoUnitPriceNOTAutoPreviousStockArray[name] = int(previousStockCountPageTwo[previousStockCountPageTwo['物品'] == pageTwoUnitPriceNOTAuto['物品'].values[name]]['现有数量'].values[0])
                                                    except IndexError:
                                                        pageTwoUnitPriceNOTAutoPreviousStockArray[name] = -404
                                                    except ValueError:
                                                        pageTwoUnitPriceNOTAutoPreviousStockArray[name] = -406

                                                pageTwoUnitPriceNOTAuto['上月存货'] = pageTwoUnitPriceNOTAutoPreviousStockArray
                                                pageTwoUnitPriceNOTAuto['上月存货'] = pageTwoUnitPriceNOTAuto['上月存货'].astype(int)

                                                pageTwoUnitPriceNOTAuto['本月使用量'] = pageTwoUnitPriceNOTAuto['上月存货'] + pageTwoUnitPriceNOTAuto['进货数量'] - pageTwoUnitPriceNOTAuto['现有数量']
                                                pageTwoUnitPriceNOTAuto['编号'] = np.arange(0, len(pageTwoUnitPriceNOTAuto))
                                                pageTwoUnitPriceNOTAuto = pageTwoUnitPriceNOTAuto[['编号','物品', '单位', '单价', "上月存货",'进货数量','本月使用量', '现有数量']]

                                                pageTwoDf = pd.concat([pageTwoUnitPriceAuto,pageTwoUnitPriceNOTAuto], ignore_index=True, axis=0)
                                                pageTwoDf['编号'] = np.arange(1, len(pageTwoDf)+1)

                                                pageTwoDfPriceRestructArray = [' ' if x is np.nan else '${}'.format(format(float(x), '.2f')) for x in pageTwoDf['单价'].values]
                                                pageTwoDf['单价'] = pageTwoDfPriceRestructArray

                                                pbar3.update(33)

                                                pageThreeUnitPrice = get_inv_df(formURL, "pageThreeUnitPrice")

                                                pageThreeUnitPriceAuto = pageThreeUnitPrice.copy()
                                                pageThreeUnitPriceAutoDropIndex = pageThreeUnitPrice[pageThreeUnitPrice['物品'] == 'BREAKLINE'].index[0]
                                                pageThreeUnitPriceAuto.drop(np.arange(pageThreeUnitPriceAutoDropIndex, len(pageThreeUnitPrice)), axis=0, inplace=True)

                                                drink_db["DATE"] = pd.to_datetime(drink_db["DATE"])
                                                drink_db = drink_db[drink_db["DATE"].isin(dateRangeFilter)]

                                                drinkInventoryFileStockInDropTitle = []

                                                for title in drink_db.columns:
                                                    if title not in ["DATE"]:
                                                        if '进' not in title:
                                                            drinkInventoryFileStockInDropTitle += [title]

                                                drinkInventoryFileStockIn = drink_db.copy()
                                                drinkInventoryFileStockIn.drop(drinkInventoryFileStockInDropTitle, axis=1, inplace=True)

                                                drinkInventoryFileStockIn = drink_db.copy()
                                                drinkInventoryFileStockIn.drop(drinkInventoryFileStockInDropTitle, axis=1, inplace=True)

                                                if drinkInventoryFileStockIn.empty:
                                                    drinkInventoryFileStockInSums = np.repeat(-404, len(pageThreeUnitPriceAuto))
                                                    print("酒水库存在数据库里没有入库信息，将以数值-404呈现")
                                                else:
                                                    drinkInventoryFileStockInSums = np.repeat(None, len(pageThreeUnitPriceAuto))

                                                    for name in range(len(pageThreeUnitPriceAuto)):
                                                        drinkInventoryFileStockInSums[name] = int(drinkInventoryFileStockIn[pageThreeUnitPriceAuto['物品'].values[name]+'进'].sum())


                                                pageThreeUnitPriceAuto['进货数量'] = drinkInventoryFileStockInSums

                                                drinkInventoryFileOutDropTitle = []

                                                for title in drink_db.columns:
                                                    if title not in ["DATE"]:
                                                        if '出' not in title:
                                                            drinkInventoryFileOutDropTitle += [title]

                                                drinkInventoryFileOut = drink_db.copy()
                                                drinkInventoryFileOut.drop(drinkInventoryFileOutDropTitle, axis=1, inplace=True)

                                                if drinkInventoryFileOut.empty:
                                                    drinkInventoryFileOutSums = np.repeat(-404, len(pageThreeUnitPriceAuto))
                                                    print("酒水库存在数据库里没有出库信息，将以数值-404呈现")
                                                else:
                                                    drinkInventoryFileOutSums = np.repeat(None, len(pageThreeUnitPriceAuto))

                                                    for name in range(len(pageThreeUnitPriceAuto)):
                                                        drinkInventoryFileOutSums[name] = int(drinkInventoryFileOut[pageThreeUnitPriceAuto['物品'].values[name]+'出'].sum())

                                                pageThreeUnitPriceAuto['本月使用量'] = drinkInventoryFileOutSums

                                                drinkInventoryCurrentDropTitle = []

                                                for title in drink_db.columns:
                                                    if title not in ['DATE']:
                                                        if '实结存' not in title:
                                                            drinkInventoryCurrentDropTitle += [title]

                                                drinkInventoryCurrent = drink_db.copy()
                                                drinkInventoryCurrent.drop(drinkInventoryCurrentDropTitle, axis=1, inplace=True)

                                                if drinkInventoryCurrent.empty:
                                                    drinkInventoryCurrentSums = np.repeat(-404, len(pageThreeUnitPriceAuto))
                                                    print("酒水库存在数据库里没有完整信息，将以数值-404呈现")
                                                else:
                                                    drinkInventoryCurrent["DATE"] = pd.to_datetime(drinkInventoryCurrent["DATE"])
                                                    maxDateFordrinkInventoryCurrent = drinkInventoryCurrent["DATE"].max()

                                                    drinkInventoryCurrent = drinkInventoryCurrent[drinkInventoryCurrent["DATE"] == maxDateFordrinkInventoryCurrent]

                                                    drinkInventoryCurrentSums = np.repeat(None, len(pageThreeUnitPriceAuto))
                                                    for name in range(len(pageThreeUnitPriceAuto)):
                                                        drinkInventoryCurrentSums[name] = int(drinkInventoryCurrent[pageThreeUnitPriceAuto['物品'].values[name]+'实结存'].sum())

                                                pageThreeUnitPriceAuto['现有数量'] = drinkInventoryCurrentSums

                                                if len(pageThreeUnitPriceAuto) == len(drink_stock_df):
                                                    nameReplace = np.repeat(None, len(pageThreeUnitPriceAuto))
                                                    for index in range(len(drink_stock_df)):
                                                        nameReplace[index] = drink_stock_df.iloc[index, 7]

                                                    pageThreeUnitPriceAuto["物品"] = nameReplace

                                                else:
                                                    print("Critical Error, drink names have not changed for stock count")

                                                pageThreeUnitPriceAuto["ACTIVE STATUS"] = pageThreeUnitPriceAuto["ACTIVE STATUS"].astype(int)
                                                pageThreeUnitPriceAuto = pageThreeUnitPriceAuto[pageThreeUnitPriceAuto["ACTIVE STATUS"] == 1]

                                                previousStockCountPageThree = pd.read_excel('{}/{}/{}/{}'.format(os.getcwd(), stock_count_foldername, stock_count_excel_foldername, previousStockCountFilename),sheet_name="Sheet3")
                                                previousStockCountPageThree.drop(['编号', '单位', '单价', '上月存货', '进货数量', '本月使用量','备注'], axis=1, inplace=True)

                                                pageThreeUnitPriceAutoPreviousStockCountArray = np.repeat(None, len(pageThreeUnitPriceAuto))

                                                for name in range(len(pageThreeUnitPriceAuto)):
                                                    try:
                                                        pageThreeUnitPriceAutoPreviousStockCountArray[name] = int(previousStockCountPageThree[previousStockCountPageThree['物品'] == pageThreeUnitPriceAuto['物品'].values[name]]['现有数量'].values[0])
                                                    except ValueError:
                                                        pageThreeUnitPriceAutoPreviousStockCountArray[name] = -406
                                                    except IndexError:
                                                        pageThreeUnitPriceAutoPreviousStockCountArray[name] = -404


                                                pageThreeUnitPriceAuto['上月存货'] = pageThreeUnitPriceAutoPreviousStockCountArray
                                                pageThreeUnitPriceAuto['编号'] = np.arange(1, len(pageThreeUnitPriceAuto)+1)
                                                pageThreeUnitPriceAuto = pageThreeUnitPriceAuto[['编号', '物品', '单位', '单价', '上月存货', '进货数量', '本月使用量', '现有数量']]

                                                pageThreeUnitPriceNOTAuto = pageThreeUnitPrice.copy()
                                                pageThreeUnitPriceNOTAuto.drop(np.arange(0, pageThreeUnitPriceAutoDropIndex+1), axis=0, inplace=True)

                                                pageThreeUnitPriceNOTAutoPreviousStockCountArray = np.repeat(None, len(pageThreeUnitPriceNOTAuto))

                                                for name in range(len(pageThreeUnitPriceNOTAuto)):
                                                    try:
                                                        pageThreeUnitPriceNOTAutoPreviousStockCountArray[name] = int(previousStockCountPageThree[previousStockCountPageThree['物品'] == pageThreeUnitPriceNOTAuto['物品'].values[name]]['现有数量'].values[0])
                                                    except IndexError:
                                                        pageThreeUnitPriceNOTAutoPreviousStockCountArray[name] = -404
                                                    except ValueError:
                                                        pageThreeUnitPriceNOTAutoPreviousStockCountArray[name] = -406

                                                pageThreeUnitPriceNOTAuto['上月存货'] = pageThreeUnitPriceNOTAutoPreviousStockCountArray

                                                buyInStockDfFilteredGroupedRemoveUnitsForPageThreeUnitPriceNOTAuto = buyInStockDfFilteredGroupedRemoveUnits.copy()
                                                buyInStockDfFilteredGroupedRemoveUnitsForPageThreeUnitPriceNOTAutoDropIndex = []

                                                for name in range(len(buyInStockDfFilteredGroupedRemoveUnitsForPageThreeUnitPriceNOTAuto)):
                                                    if buyInStockDfFilteredGroupedRemoveUnitsForPageThreeUnitPriceNOTAuto['物品(单位)'].values[name] not in pageThreeUnitPriceNOTAuto['物品'].values:
                                                        buyInStockDfFilteredGroupedRemoveUnitsForPageThreeUnitPriceNOTAutoDropIndex += [name]
                                                buyInStockDfFilteredGroupedRemoveUnitsForPageThreeUnitPriceNOTAuto.drop( buyInStockDfFilteredGroupedRemoveUnitsForPageThreeUnitPriceNOTAutoDropIndex, axis=0, inplace=True)
                                                buyInStockDfFilteredGroupedRemoveUnitsForPageThreeUnitPriceNOTAuto.reset_index(inplace=True)
                                                buyInStockDfFilteredGroupedRemoveUnitsForPageThreeUnitPriceNOTAuto.drop("index", axis=1, inplace=True)
                                                buyInStockDfFilteredGroupedRemoveUnitsForPageThreeUnitPriceNOTAuto.columns = ['物品', '进货数量']

                                                pageThreeUnitPriceNOTAuto = pd.merge(right=buyInStockDfFilteredGroupedRemoveUnitsForPageThreeUnitPriceNOTAuto, left=pageThreeUnitPriceNOTAuto, how='outer')

                                                buyInStockDfFilteredGroupedRemoveUnitsForPageThreeUnitPriceNOTAutoTurnZero = [0 if str(x) == 'nan' else x for x in pageThreeUnitPriceNOTAuto['进货数量'].values]
                                                pageThreeUnitPriceNOTAuto['进货数量'] = buyInStockDfFilteredGroupedRemoveUnitsForPageThreeUnitPriceNOTAutoTurnZero

                                                pageThreeUnitPriceNOTAutoCurrentArray = np.repeat(None,len(pageThreeUnitPriceNOTAuto))

                                                for name in range(len(pageThreeUnitPriceNOTAuto)):
                                                    pageThreeUnitPriceNOTAutoCurrentArray[name] = stockCountDfFilteredRemoveUnits[pageThreeUnitPriceNOTAuto['物品'].values[name]]

                                                pageThreeUnitPriceNOTAuto['现有数量'] = pageThreeUnitPriceNOTAutoCurrentArray
                                                pageThreeUnitPriceNOTAuto['现有数量'] = pageThreeUnitPriceNOTAuto['现有数量'].astype(int)
                                                pageThreeUnitPriceNOTAuto['上月存货'] = pageThreeUnitPriceNOTAuto['上月存货'].astype(int)
                                                pageThreeUnitPriceNOTAuto['进货数量'] = pageThreeUnitPriceNOTAuto['进货数量'].astype(int)

                                                pageThreeUnitPriceNOTAuto['本月使用量'] = pageThreeUnitPriceNOTAuto['上月存货'] + pageThreeUnitPriceNOTAuto['进货数量'] - pageThreeUnitPriceNOTAuto['现有数量']
                                                pageThreeUnitPriceNOTAuto['编号'] = np.arange(1, len(pageThreeUnitPriceNOTAuto)+1)
                                                pageThreeUnitPriceNOTAuto = pageThreeUnitPriceNOTAuto[['编号', '物品', '单位', '单价', '上月存货', '进货数量', '本月使用量', '现有数量']]

                                                pageThreeDf = pd.concat([pageThreeUnitPriceAuto, pageThreeUnitPriceNOTAuto], axis=0, ignore_index=True)
                                                pageThreeDf['编号'] = np.arange(1, len(pageThreeDf)+1)

                                                pageThreeDfPriceRestruct = np.repeat(None, len(pageThreeDf))

                                                for price in range(len(pageThreeDf['单价'].values.astype(str))):
                                                    if pageThreeDf['单价'].values.astype(str)[price] == ' ':
                                                        pageThreeDfPriceRestruct[price] = ' '
                                                    else:
                                                        pageThreeDfPriceRestruct[price] = '${}'.format(format(float(pageThreeDf['单价'].values.astype(str)[price]), '.2f'))

                                                pageThreeDf['单价'] = pageThreeDfPriceRestruct

                                                pbar3.update(33)

                                                pageOneDf['备注'] = np.nan
                                                pageTwoDf['备注'] = np.nan
                                                pageThreeDf['备注'] = np.nan

                                                pbar3.set_description("完成")
                                                pbar3.update(1)

                                                prtdf(pageOneDf)
                                                prtdf(pageTwoDf)
                                                prtdf(pageThreeDf)

                                                current_stock_filename = "{}年{}月{}盘点详情.xlsx".format(stockStartDate.year, stockStartDate.month, storeOutlet.strip().capitalize())

                                                if os.path.exists("{}/{}/{}/{}".format(os.getcwd(), stock_count_foldername, stock_count_excel_foldername, current_stock_filename)):
                                                    print("本次的盘点文件'{}'已经存在，是否覆盖?".format(current_stock_filename))
                                                    print("你如果已经对该文件做了修改，你覆盖之后将会丢失所有你修改过的数据")

                                                    saveActions = option_num(['保存', '丢弃'])
                                                    time.sleep(0.25)
                                                    userInputSix = option_limit(saveActions, input("在这里输入>>>: "))
                                                    if userInputSix == 0:
                                                        with pd.ExcelWriter("{}/{}/{}/{}".format(os.getcwd(), stock_count_foldername, stock_count_excel_foldername, current_stock_filename)) as writer:
                                                            pageOneDf.to_excel(writer, sheet_name='Sheet1', index=False, header=True)
                                                            pageTwoDf.to_excel(writer, sheet_name='Sheet2', index=False, header=True)
                                                            pageThreeDf.to_excel(writer, sheet_name='Sheet3', index=False, header=True)
                                                        print("文件已覆盖")
                                                    elif userInputSix == 1:
                                                        print("文件已丢弃")
                                                else:
                                                    with pd.ExcelWriter("{}/{}/{}/{}".format(os.getcwd(), stock_count_foldername, stock_count_excel_foldername, current_stock_filename)) as writer:
                                                        pageOneDf.to_excel(writer, sheet_name='Sheet1', index=False, header=True)
                                                        pageTwoDf.to_excel(writer, sheet_name='Sheet2', index=False, header=True)
                                                        pageThreeDf.to_excel(writer, sheet_name='Sheet3', index=False, header=True)
                                                    print("文件已自动保存")
                                                    print("前往盘点详情excel的文件夹内查看, 文件名为'{}'".format(current_stock_filename))

                                        else:
                                            print("无法获取数据库信息, 盘点无法继续")

                            elif userInputFive == 1:
                                fileShown = os.listdir("{}/{}/{}/".format(os.getcwd(), stock_count_foldername, stock_count_excel_foldername))

                                fileToShow = []

                                for file in fileShown:
                                    if file.startswith("~"):
                                        pass
                                    elif file.startswith("."):
                                        pass
                                    else:
                                        if file.endswith(".xlsx"):
                                            fileToShow += [file]
                                        else:
                                            pass

                                fileToShow += ["返回上一菜单"]

                                selectFileToViewOptions = option_num(fileToShow)
                                time.sleep(0.25)
                                userInputSeven = option_limit(selectFileToViewOptions, input("在这里输入>>>: "))

                                if userInputSeven == len(fileToShow)-1 :
                                    pass
                                else:
                                    print(fileToShow[userInputSeven])
                                    prtdf(pd.read_excel("{}/{}/{}/{}".format(os.getcwd(), stock_count_foldername, stock_count_excel_foldername, fileToShow[userInputSeven]), sheet_name="Sheet1"))
                                    prtdf(pd.read_excel("{}/{}/{}/{}".format(os.getcwd(), stock_count_foldername, stock_count_excel_foldername, fileToShow[userInputSeven]), sheet_name="Sheet2"))
                                    prtdf(pd.read_excel("{}/{}/{}/{}".format(os.getcwd(), stock_count_foldername, stock_count_excel_foldername, fileToShow[userInputSeven]), sheet_name="Sheet3"))

                    elif userInputOne == 2:
                        fileShown = os.listdir("{}/{}/{}/".format(os.getcwd(), stock_count_foldername, stock_count_excel_foldername))

                        fileToShow = []

                        for file in fileShown:
                            if file.startswith("~"):
                                pass
                            elif file.startswith("."):
                                pass
                            else:
                                if file.endswith(".xlsx"):
                                    fileToShow += [file]
                                else:
                                    pass

                        fileToShow += ["返回上一菜单"]

                        selectFileToViewOptions = option_num(fileToShow)
                        time.sleep(0.25)
                        userInputEight = option_limit(selectFileToViewOptions, input("在这里输入>>>: "))

                        if userInputEight == len(fileToShow)-1:
                            pass
                        else:
                            try:
                                fileNameAfterCheck = doc_filename_check(fileName=fileToShow[userInputEight])

                                if os.path.exists("{}/{}/{}".format(os.getcwd(), stock_count_foldername, songti_filename)):
                                    try:
                                        pageOneDf = pd.read_excel("{}/{}/{}/{}".format(os.getcwd(), stock_count_foldername, stock_count_excel_foldername, fileToShow[userInputEight]), sheet_name="Sheet1")
                                        pageTwoDf = pd.read_excel("{}/{}/{}/{}".format(os.getcwd(), stock_count_foldername, stock_count_excel_foldername, fileToShow[userInputEight]), sheet_name="Sheet2")
                                        pageThreeDf = pd.read_excel("{}/{}/{}/{}".format(os.getcwd(), stock_count_foldername, stock_count_excel_foldername, fileToShow[userInputEight]), sheet_name="Sheet3")

                                        print("读取字体文件中...")
                                        custom_font_path = pathlib.Path("{}/{}/{}".format(os.getcwd(), stock_count_foldername, songti_filename))
                                        borb_custom_font = borb_TrueTypeFont.true_type_font_from_file(custom_font_path)
                                        print("字体文件读取完成。")
                                        print()
                                        print("库存盘点PDF生成过程比较漫长, 请耐心等待...")
                                        print()
                                        stock_gen_pdf(customFont=borb_custom_font,
                                                      pageOneDf=pageOneDf,
                                                      pageTwoDf=pageTwoDf,
                                                      pageThreeDf=pageThreeDf,
                                                      fileNameAfterCheck=fileNameAfterCheck,
                                                      outlet=storeOutlet,
                                                      stock_count_foldername=stock_count_foldername,
                                                      stock_count_pdf_foldername=stock_count_pdf_foldername,)

                                        print()
                                        print()

                                    except ValueError:
                                        print("工作簿的文件名不对，必须是'Sheet1', 'Sheet2', 'Sheet3'且")
                                        print("Sheet1必须对应楼面盘点")
                                        print("Sheet2必须对应打包盒盘点")
                                        print("Sheet3必须对应收银盘点")
                                        print("无法继续生成PDF")
                                        print()

                                else:
                                    print("字体宋体ttf文件不存在")
                                    print("请将字体宋体ttf文件保存至路径'{}/{}/{}'".format(os.getcwd(), stock_count_foldername, songti_filename))

                            except TypeError:
                                print("选择的{}文件名有误，必须是【几几几几】年【几几】月【门店名】盘点详情.xlsx".format(fileToShow[userInputEight]))
                                print()

def point_five_round(x):
    return normal_round(normal_round(x / 0.5) * 0.5, -int(np.floor(np.log10(0.5))))

def getWeeklyDateRange(date):

    weekday = date.weekday() #0 is Mon, 6 is Sun

    parseDict = {
        0 : pd.date_range(start=date, end=(date + dt.timedelta(days=6))),
        1 : pd.date_range(start=(date - dt.timedelta(days=1)), end=(date + dt.timedelta(days=5))),
        2 : pd.date_range(start=(date - dt.timedelta(days=2)), end=(date + dt.timedelta(days=4))),
        3 : pd.date_range(start=(date - dt.timedelta(days=3)), end=(date + dt.timedelta(days=3))),
        4 : pd.date_range(start=(date - dt.timedelta(days=4)), end=(date + dt.timedelta(days=2))),
        5 : pd.date_range(start=(date - dt.timedelta(days=5)), end=(date + dt.timedelta(days=1))),
        6 : pd.date_range(start=(date - dt.timedelta(days=6)), end=date),
    }

    return parseDict[weekday]

def getLastEntry(employee_id, date, shift_database, last_entry_for):

    criteria0 = shift_database["ID"] == str(employee_id)
    criteria1 = shift_database["DATE"] < date

    filteredDf = shift_database[(criteria0)&(criteria1)]

    if filteredDf.empty:
        lastEntry = 0.0

    else:
        getMaxDate = filteredDf["DATE"].max()
        lastEntry = filteredDf[filteredDf["DATE"] == getMaxDate][str(last_entry_for)].values[0]

    return lastEntry

def calculate_AL(employee_info_df, employee_id, date):
    criteria1 = employee_info_df["ID"] == str(employee_id)

    AL_eligibility = eval(str(employee_info_df[(criteria1)]["AL Eligibility"].values[0]).strip().capitalize())

    if AL_eligibility:
        AL_end_date = pd.to_datetime(employee_info_df[(criteria1)]["AL END DATE"].values[0])
        AL_start_date = pd.to_datetime(employee_info_df[(criteria1)]["AL START DATE"].values[0])

        if date == AL_end_date:

            if AL_start_date >= AL_end_date:
                AL_value = -1000000.00

            else:
                AL_base = float(str(employee_info_df[(criteria1)]["AL DAY BASE"].values[0]))
                unpaid_months = float(str(employee_info_df[(criteria1)]["UNPAID MONTH(S) TAKEN WITHIN PERIOD OF AL CALCULATION"].values[0]))

                cut_off_day = 1
                cut_off_month = 9

                first_day = pd.to_datetime(employee_info_df[(criteria1)]["FIRST DAY DATE"].values[0])
                cut_off_date = pd.to_datetime(dt.datetime(first_day.year, cut_off_month, cut_off_day))

                if first_day < cut_off_date:
                    date_of_increment = pd.to_datetime(dt.datetime(first_day.year+1, 1, 1))

                else:
                    date_of_increment = pd.to_datetime(dt.datetime(first_day.year+2, 1,1))

                if first_day > AL_start_date:
                    AL_start_date = first_day

                completed_month_of_service = len(pd.date_range(start=first_day, end=AL_end_date, freq="MS"))

                if completed_month_of_service < 3:
                    AL_value = 0

                else:
                    if AL_start_date >= date_of_increment:
                        consider_month = len(pd.date_range(start=date_of_increment, end=AL_end_date, freq="MS"))
                        AL_days = consider_month / 12 + AL_base

                    else:
                        AL_days = AL_base


                    total_months = len(pd.date_range(start=AL_start_date, end=AL_end_date, freq="MS"))
                    AL_value = total_months / 12 * AL_days

                    AL_value = AL_value - ((AL_value/total_months)*unpaid_months)

        else:
            AL_value = 0
    else:
        AL_value = 0

    return point_five_round(AL_value)

def assignLeaves(assign_for, employee_id, date, employee_info_df, shift_df, leaves_manual_df, ph_dates_df):

    weekday = date.weekday() #0 is Mon, 6 is Sun

    criteria0 = employee_info_df["ID"] == str(employee_id)
    criteria1 = shift_df["ID"] == str(employee_id)
    criteria2 = shift_df["DATE"].isin(getWeeklyDateRange(date))
    OFF_string = "OFF"

    if str(assign_for).strip().upper() == "OIL":

        OT_string = "OT"
        OIL_string = "OIL"

        offInAWeek = float(employee_info_df[(criteria0)]["AVERAGE OFF IN A WEEK"].values[0])
        isAuto = str(employee_info_df[(criteria0)]["HOW OIL ACCUMULATE"].values[0])

        if weekday == 6 and isAuto.strip().upper() == "AUTO": #only top up OIL on a Sunday

            if len(shift_df[(criteria1)&(criteria2)]) != 7:
                #print(employee_id, "incomplete record")
                AssignValue = 0.0

            else:
                criteria3 = shift_df["SHIFT"] == OFF_string
                criteria4 = shift_df["SHIFT"].str.startswith(OFF_string+"/")
                criteria5 = shift_df["SHIFT"].str.endswith("/"+OFF_string)

                num1 = len(shift_df[(criteria1)&(criteria2)&(criteria3)])
                num2 = len(shift_df[(criteria1)&(criteria2)&(criteria4)])*0.5
                num3 = len(shift_df[(criteria1)&(criteria2)&(criteria5)])*0.5

                criteria6 = shift_df["SHIFT"] == OT_string
                criteria7 = shift_df["SHIFT"].str.startswith(OT_string+"/")
                criteria8 = shift_df["SHIFT"].str.endswith("/"+OT_string)

                num4 = len(shift_df[(criteria1)&(criteria2)&(criteria6)])
                num5 = len(shift_df[(criteria1)&(criteria2)&(criteria7)])*0.5
                num6 = len(shift_df[(criteria1)&(criteria2)&(criteria8)])*0.5

                AssignValue = offInAWeek - (num1 + num2 + num3) - (num4 + num5 + num6)

        else:
            AssignValue = 0.0

        onDateCriteria = shift_df["DATE"] == date

        OIL_criteria1 = shift_df["SHIFT"] == str(assign_for).strip().upper()
        OIL_criteria2 = shift_df["SHIFT"].str.endswith("/"+str(assign_for).strip().upper())
        OIL_criteria3 = shift_df["SHIFT"].str.startswith(str(assign_for).strip().upper()+"/")

        num7 = len(shift_df[(criteria1)&(onDateCriteria)&(OIL_criteria1)])
        num8 = len(shift_df[(criteria1)&(onDateCriteria)&(OIL_criteria2)])*0.5
        num9 = len(shift_df[(criteria1)&(onDateCriteria)&(OIL_criteria3)])*0.5

        AssignValue -= (num7 + num8 + num9)

    elif str(assign_for).strip().upper() == "PH":
        ph_eligibility = eval(str(employee_info_df[criteria0]["PH Eligibility"].values[0]).capitalize())

        if ph_eligibility:
            if date in ph_dates_df["PH DATE"].values:
                criteria3 = shift_df["SHIFT"] == OFF_string
                criteria4 = shift_df["DATE"] == date

                criteria5 = shift_df["SHIFT"].str.startswith(OFF_string+"/")
                criteria6 = shift_df["SHIFT"].str.endswith("/"+OFF_string)

                num1 = len(shift_df[(criteria1)&(criteria4)&(criteria3)])
                num2 = len(shift_df[(criteria1)&(criteria4)&(criteria5)])*0.5
                num3 = len(shift_df[(criteria1)&(criteria4)&(criteria6)])*0.5

                AssignValue = 1.0 - (num1 + num2 + num3)

            else:
                AssignValue = 0.0

            criteria4 = shift_df["DATE"] == date

            criteria3 = shift_df["SHIFT"] == str(assign_for).strip().upper()
            criteria5 = shift_df["SHIFT"].str.startswith(str(assign_for).strip().upper()+"/")
            criteria6 = shift_df["SHIFT"].str.endswith("/"+str(assign_for).strip().upper())

            num4 = len(shift_df[(criteria1)&(criteria4)&(criteria3)])
            num5 = len(shift_df[(criteria1)&(criteria4)&(criteria5)])*0.5
            num6 = len(shift_df[(criteria1)&(criteria4)&(criteria6)])*0.5

            AssignValue -= (num4 + num5 + num6)

        else:
            AssignValue = 0.0
    elif str(assign_for).strip().upper() == "AL":
        AssignValue = calculate_AL(employee_info_df=employee_info_df, employee_id=employee_id, date=date)

        criteria3 = shift_df["SHIFT"] == str(assign_for).strip().upper()
        criteria4 = shift_df["DATE"] == date
        criteria5 = shift_df["SHIFT"].str.startswith(str(assign_for).strip().upper()+"/")
        criteria6 = shift_df["SHIFT"].str.endswith("/"+str(assign_for).strip().upper())

        num1 = len(shift_df[(criteria1)&(criteria4)&(criteria3)])
        num2 = len(shift_df[(criteria1)&(criteria4)&(criteria5)])*0.5
        num3 = len(shift_df[(criteria1)&(criteria4)&(criteria6)])*0.5

        total_number = num1 + num2 + num3

        AssignValue -= total_number

    else:
        filter1 = shift_df["DATE"] == date

        filter2 = shift_df["SHIFT"] == str(assign_for).strip().upper()
        filter3 = shift_df["SHIFT"].str.startswith(str(assign_for).strip().upper()+"/")
        filter4 = shift_df["SHIFT"].str.endswith("/"+str(assign_for).strip().upper())

        number1 = len(shift_df[(criteria1)&(filter1)&(filter2)])
        number2 = len(shift_df[(criteria1)&(filter1)&(filter3)])*0.5
        number3 = len(shift_df[(criteria1)&(filter1)&(filter4)])*0.5

        AssignValue = 0.0 - (number1 + number2 + number3)


    criteria10 = leaves_manual_df["ID"] == str(employee_id)
    criteria11 = leaves_manual_df["LEAVES TYPE"] == str(assign_for).strip().upper()
    criteria12 = leaves_manual_df["DATE"] == date

    get_manual = leaves_manual_df[(criteria10)&(criteria11)&(criteria12)]
    if get_manual.empty:
        get_manual = 0.0
    else:
        get_manual = float(get_manual["VALUE"].values[0])

    return point_five_round(AssignValue + get_manual)

def writeEntry(date, employee_id, shift, OILPlus, PHPlus, ALPlus, CCLPlus, remarks, shift_database, employee_info_df, ph_dates_df):

    DATE = date
    DAY_NAME = date.strftime("%A")
    ID = str(employee_id)

    FOR_VLOOKUP = "{}${}".format(ID, DATE.strftime("%Y-%m-%d"))

    NAME = employee_info_df[employee_info_df["ID"] == str(employee_id)]["NAME"].values[0]
    SHIFT = shift
    OIL=float(getLastEntry(employee_id=employee_id, date=date, shift_database=shift_database, last_entry_for="OIL")) + float(OILPlus)
    PH=float(getLastEntry(employee_id=employee_id, date=date, shift_database=shift_database, last_entry_for="PH"))+float(PHPlus)
    AL=float(getLastEntry(employee_id=employee_id, date=date, shift_database=shift_database, last_entry_for="AL"))+float(ALPlus)
    CCL=float(getLastEntry(employee_id=employee_id, date=date, shift_database=shift_database, last_entry_for="CCL"))+float(CCLPlus)

    IS_PH = not (ph_dates_df[ph_dates_df["PH DATE"] == date].empty)

    REMARKS = remarks

    TIME_LOG = pd.to_datetime(dt.datetime.now())

    col = shift_database.columns
    DataFrame = pd.DataFrame({
        col[0] : [FOR_VLOOKUP],
        col[1] : [DATE],
        col[2] : [DAY_NAME],
        col[3] : [ID],
        col[4] : [NAME],
        col[5] : [SHIFT],
        col[6] : [OIL],
        col[7] : [PH],
        col[8] : [AL],
        col[9] : [CCL],
        col[10] : [str(IS_PH).strip().upper()],
        col[11] : [REMARKS],
        col[12] : [TIME_LOG],
    })

    return DataFrame

def payslip_time(shift_database, start_date, end_date, employee_info_df):
    desired_df = pd.DataFrame(columns=["ID", "NAME", "NP", "MC", "OT", "OFF DAYS",
                                       "OIL TAKEN", "PH TAKEN", "AL TAKEN", "CCL TAKEN",
                                       "TS", "BD", "BG", "A", "B", "Office"])

    c1 = shift_database["DATE"] >= start_date
    c2 = shift_database["DATE"] <= end_date

    filteredDatabase = shift_database[(c1)&(c2)]
    concatList = [desired_df]

    for i in range(len(employee_info_df)):
        if str(employee_info_df.iloc[i, 7]).strip().upper() == "TRUE":
            value = {}

            for item in ["NP", "MC", "OT", "OFF", "OIL", "PH", "AL", "CCL", "TS", "BD", "BG", "A", "B", "Office"]:

                c3 = filteredDatabase["SHIFT"] == item
                c4 = filteredDatabase["SHIFT"].str.startswith(item + "/")
                c5 = filteredDatabase["SHIFT"].str.endswith("/" + item)
                c6 = filteredDatabase["ID"] == str(employee_info_df.iloc[i, 0])

                v = float(len(filteredDatabase[(c3)&(c6)]))
                v += float(len(filteredDatabase[(c4)&(c6)]))*0.5
                v += float(len(filteredDatabase[(c5)&(c6)]))*0.5

                value.update({ item : point_five_round(v) })

            df = pd.DataFrame({
                "ID" : [str(employee_info_df.iloc[i, 0])],
                "NAME" : [str(employee_info_df.iloc[i, 1])],
                "NP" : [value["NP"]],
                "MC" : [value["MC"]],
                "OT" : [value["OT"]],
                "OFF DAYS" : [value["OFF"]],
                "OIL TAKEN" : [value["OIL"]],
                "PH TAKEN" : [value["PH"]],
                "AL TAKEN" : [value["AL"]],
                "CCL TAKEN" : [value["CCL"]],
                "TS" : [value["TS"]],
                "BD" : [value["BD"]],
                "BG" : [value["BG"]],
                "A" : [value["A"]],
                "B" : [value["B"]],
                "Office" : [value["Office"]],
            })

            concatList += [df]

        else:
            continue

    concat = pd.concat(concatList)

    return concat

def payslip_time_statement(concat_df, employee_info_df, start_date, end_date):

    normal_string = "正常"
    work_string = "工作{}天(包括带薪假)"
    unpaid_string = "无薪假{}天"
    mc_string = "病假{}天"
    full_attendance_string = "{}满勤{}"
    ot_string = "**有加班{}天"

    unpaid_index = concat_df.columns.get_loc(key="NP")
    mc_index = concat_df.columns.get_loc(key="MC")
    ot_index = concat_df.columns.get_loc(key="OT")

    period_days = len(pd.date_range(start=start_date, end=end_date))

    total_working_index = [concat_df.columns.get_loc(key="OIL TAKEN"),
                           concat_df.columns.get_loc(key="PH TAKEN"),
                           concat_df.columns.get_loc(key="AL TAKEN"),
                           concat_df.columns.get_loc(key="CCL TAKEN"),
                           concat_df.columns.get_loc(key="TS"),
                           concat_df.columns.get_loc(key="BD"),
                           concat_df.columns.get_loc(key="BG"),
                           concat_df.columns.get_loc(key="A"),
                           concat_df.columns.get_loc(key="B"),
                           concat_df.columns.get_loc(key="Office"),
                          ]

    paid_leaves_index = [concat_df.columns.get_loc(key="OIL TAKEN"),
                         concat_df.columns.get_loc(key="PH TAKEN"),
                         concat_df.columns.get_loc(key="AL TAKEN"),
                         concat_df.columns.get_loc(key="CCL TAKEN")]

    total_record_index = [concat_df.columns.get_loc(key="NP"),
                          concat_df.columns.get_loc(key="MC"),
                          concat_df.columns.get_loc(key="OT"),
                          concat_df.columns.get_loc(key="OFF DAYS"),
                          concat_df.columns.get_loc(key="OIL TAKEN"),
                          concat_df.columns.get_loc(key="PH TAKEN"),
                          concat_df.columns.get_loc(key="AL TAKEN"),
                          concat_df.columns.get_loc(key="CCL TAKEN"),
                          concat_df.columns.get_loc(key="TS"),
                          concat_df.columns.get_loc(key="BD"),
                          concat_df.columns.get_loc(key="BG"),
                          concat_df.columns.get_loc(key="A"),
                          concat_df.columns.get_loc(key="B"),
                          concat_df.columns.get_loc(key="Office"),
                         ]

    statement_dict = {}
    for index in range(len(concat_df)):
        unpaid_days = float(concat_df.iloc[index, unpaid_index])
        mc_days = float(concat_df.iloc[index, mc_index])

        total_record = 0
        for t in total_record_index:
            total_record += float(concat_df.iloc[index, t])

        criteria1 = total_record == period_days
        criteria2 = unpaid_days + mc_days == 0

        satisfy1 = total_record == period_days
        satisfy2 = unpaid_days + mc_days == 0
        satisfy3 = (mc_days > 0) and (unpaid_days == 0)

        if (satisfy1) and (satisfy2):
            statement_dict.update({ str(concat_df.iloc[index, 1]) : normal_string })

        else:
            if (satisfy1) and (satisfy3):
                statement_dict.update({ str(concat_df.iloc[index, 1]) : normal_string + "(" + mc_string.format(point_five_round(mc_days)) + ")"})

            else:
                work_days  = 0
                for w in total_working_index:
                    work_days += float(concat_df.iloc[index, w])

                statement_dict.update({ str(concat_df.iloc[index, 1]) : work_string.format(point_five_round(work_days)) })

                if unpaid_days > 0:
                    if str(concat_df.iloc[index, 1]) in statement_dict:
                        statement_dict[str(concat_df.iloc[index, 1])] += ", " + unpaid_string.format(point_five_round(unpaid_days))
                    else:
                        statement_dict.update({ str(concat_df.iloc[index, 1]) : unpaid_string.format(point_five_round(unpaid_days))})

                if mc_days > 0:
                    if str(concat_df.iloc[index, 1]) in statement_dict:
                        statement_dict[str(concat_df.iloc[index, 1])] += ", " + mc_string.format(point_five_round(mc_days))
                    else:
                        statement_dict.update({ str(concat_df.iloc[index, 1]) : mc_string.format(point_five_round(mc_days))})

        full_attendance_consider = eval(str(employee_info_df[employee_info_df["ID"] == str(concat_df.iloc[index, 0])]["FULL ATTENDANCE CONSIDER"].values[0]).strip().capitalize())

        if (full_attendance_consider) and (satisfy1):
            is_late = eval(str(employee_info_df[employee_info_df["ID"] == str(concat_df.iloc[index, 0])]["IS LATE WITHIN PERIOD OF PAYSLIP TIME CALCULATION"].values[0]).strip().capitalize())
            paid_leaves_cut_off  = float(str(employee_info_df[employee_info_df["ID"] == str(concat_df.iloc[index, 0])]["CUT OFF DAYS OF PAID LEAVES ALLOWABLE FOR FULL ATTENDANCE WITHIN PERIOD OF PAYSLIP TIME CALCULATION"].values[0]).strip())

            paid_days = 0
            for p in paid_leaves_index:
                paid_days += float(concat_df.iloc[index, p])

            satisfy4 = (unpaid_days == 0) and (mc_days == 0)

            if (satisfy4):
                if paid_days >= paid_leaves_cut_off and is_late:
                    statement_dict[str(concat_df.iloc[index, 1])] += ", " + full_attendance_string.format("无", "(迟到, 带薪假超过{}天)".format(point_five_round(paid_leaves_cut_off)))

                elif paid_days >= paid_leaves_cut_off and not is_late:
                    statement_dict[str(concat_df.iloc[index, 1])] += ", " + full_attendance_string.format("无", "(带薪假超过{}天)".format(point_five_round(paid_leaves_cut_off)))

                elif paid_days < paid_leaves_cut_off and is_late:
                    statement_dict[str(concat_df.iloc[index, 1])] += ", " + full_attendance_string.format("无", "(迟到)")

                elif paid_days < paid_leaves_cut_off and not is_late:
                    statement_dict[str(concat_df.iloc[index, 1])] += ", " + full_attendance_string.format("", "")
            else:
                statement_dict[str(concat_df.iloc[index, 1])] += ", " + full_attendance_string.format("无", "")

        else:
            if not(satisfy1):
                pass
            else:
                statement_dict[str(concat_df.iloc[index, 1])] += ", " + full_attendance_string.format("无", "")


        ot_days = float(concat_df.iloc[index, ot_index])

        if ot_days > 0:
            statement_dict[str(concat_df.iloc[index, 1])] += ", " + ot_string.format(point_five_round(ot_days))

    return statement_dict

def retrieve_shift_db(database_url, shift_db_sheetname, local_db_filename, backup_foldername):
    #load shift database
    shift_db_columns = ["FOR VLOOKUP", "DATE", "DAY NAME", "ID", "NAME", "SHIFT", "OIL", "PH", "AL", "CCL", "IS PH", "REMARKS", "TIME LOG"]

    try:
        shift_db = google_auth.open_by_url(database_url)
        shift_db_sheetname_index = shift_db.worksheet(property="title", value=shift_db_sheetname).index
        shift_database_online = shift_db[shift_db_sheetname_index].get_as_df()

    except Exception as e:
        print()
        print("从云端读取排班数据库失败，错误描述如下：")
        print(e)
        print()
        shift_database_online = None

    try:
        shift_database_local = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_db_filename), sheet_name=shift_db_sheetname)

    except Exception as e:
        print()
        print("从本地读取排班数据库失败，错误描述如下：")
        print(e)
        print()
        shift_database_local = None

    if isinstance(shift_database_online, pd.DataFrame) or isinstance(shift_database_local, pd.DataFrame):
        if isinstance(shift_database_online, pd.DataFrame):
            if shift_database_online.empty:
                shift_database_online = pd.DataFrame(columns=shift_db_columns)
                shift_db_online_max_date = None
            else:
                shift_database_online["DATE"] = pd.to_datetime(shift_database_online["DATE"])
                shift_database_online["TIME LOG"] = pd.to_datetime(shift_database_online["TIME LOG"])
                shift_db_online_max_date = shift_database_online["DATE"].max()

        else:
            shift_db_online_max_date = None

        if isinstance(shift_database_local, pd.DataFrame):
            if shift_database_local.empty:
                shift_database_local = pd.DataFrame(columns=shift_db_columns)
                shift_db_local_max_date = None
            else:
                shift_database_local["DATE"] = pd.to_datetime(shift_database_local["DATE"])
                shift_database_local["TIME LOG"] = pd.to_datetime(shift_database_local["TIME LOG"])
                shift_db_local_max_date = shift_database_local["DATE"].max()

        else:
            shift_db_local_max_date = None


        if (shift_db_online_max_date == None) and (shift_db_local_max_date != None):
            shift_database = shift_database_local.copy()

        elif (shift_db_local_max_date == None) and (shift_db_online_max_date != None):
            shift_database = shift_database_online.copy()

        elif (shift_db_local_max_date != None) and (shift_db_online_max_date != None):
            if shift_db_local_max_date >= shift_db_online_max_date:
                shift_database = shift_database_local.copy()
            else:
                shift_database = shift_database_online.copy()
        else:
            shift_database = pd.DataFrame(columns=shift_db_columns)
    else:
        shift_database = None

    return shift_database

def work_schedule_main(google_auth, db_setting_url, constants_sheetname, serialized_rule_filename, backup_foldername, database_url, payslip_on_duty):
    google_auth = google_auth
    db_setting_url = db_setting_url
    constants_sheetname = constants_sheetname
    serialized_rule_filename = serialized_rule_filename
    backup_foldername = backup_foldername
    database_url = database_url
    payslip_on_duty = payslip_on_duty

    res = pyfiglet.figlet_format("Schedule")
    print(res)
    time.sleep(0.15)

    on_net = on_internet()

    if not on_net:
        print()
        print()
        print("无网络连接。")
        print("排班程序需要全程连接网络来完成。")

    else:
        fernet_key = get_key()

        if fernet_key == 0:
            print("安全密钥错误! ")
        else:
            print("初始化...")
            shiftOutlet = get_outlet()

            k_dict = get_k_dictionary(google_auth, db_setting_url, constants_sheetname, serialized_rule_filename, shiftOutlet, fernet_key, backup_foldername)
            shiftURL = fernet_decrypt(k_dict["shift_url"], fernet_key)

            database_url = fernet_decrypt(database_url, fernet_key)

            local_database_filename = str(k_dict["local_database_filename"])

            drink_db_sheetname = str(k_dict["drink_database_sheetname"])
            takeaway_box_db_sheetname = str(k_dict["takeaway_box_database_sheetname"])
            finance_db_sheetname = str(k_dict["financial_database_sheetname"])
            promotion_db_sheetname = str(k_dict["promotion_database_sheetname"])

            shift_db_sheetname = str(k_dict["shift_database_sheetname"])
            shift_sheetname = str(k_dict["shift_sheetname"])
            shift_leaves_manual_sheetname = str(k_dict["leaves_manual_sheetname"])

            tabox_stock_sheetname = str(k_dict["takeaway_box_sheetname"])
            drink_stock_sheetname = str(k_dict["drink_stock_sheetname"])
            rcv_sheetname = str(k_dict["receivers_sheetname"])

            shift_database = retrieve_shift_db(database_url, shift_db_sheetname, local_database_filename, backup_foldername)

            if isinstance(shift_database, pd.DataFrame):
                userInputOne = 0
                while userInputOne != 5:
                    with tqdm(total=100) as pbar1:

                        pbar1.set_description("处理数据库")
                        shift_database["DATE"] = pd.to_datetime(shift_database["DATE"])
                        shift_database["TIME LOG"] = pd.to_datetime(shift_database["TIME LOG"])
                        shift_database["ID"] = shift_database["ID"].astype(int)
                        shift_database["ID"] = shift_database["ID"].astype(str)
                        pbar1.update(16)

                        pbar1.set_description("读取排班全记录")
                        #df = pd.read_html(shiftURL+"htmlview", encoding="utf-8")
                        df = google_auth.open_by_url(shiftURL)
                        pbar1.update(16)

                        pbar1.set_description("读取排班记录")
                        #shift_df = df[0]
                        #parseGoogleHTMLSheet(shift_df)
                        shift_df = df[0].get_as_df()
                        shift_df = shift_df[shift_df.iloc[:,0].astype(str).str.len() > 0]
                        shift_df["DATE"] = pd.to_datetime(shift_df["DATE"])
                        shift_df["SHIFT"] = shift_df["SHIFT"].astype(str)
                        shift_df["ID"] = shift_df["ID"].astype(int)
                        shift_df["ID"] = shift_df["ID"].astype(str)
                        pbar1.update(16)

                        pbar1.set_description("读取员工资料")
                        #employee_info = df[2]
                        #parseGoogleHTMLSheet(employee_info)
                        employee_info = df[2].get_as_df()
                        employee_info = employee_info[employee_info.iloc[:,0].astype(str).str.len() > 0]
                        employee_info["ID"] = employee_info["ID"].astype(int)
                        employee_info["ID"] = employee_info["ID"].astype(str)
                        employee_info["FIRST DAY DATE"] = pd.to_datetime(employee_info["FIRST DAY DATE"])
                        employee_info["AL START DATE"] = pd.to_datetime(employee_info["AL START DATE"])
                        employee_info["AL END DATE"] = pd.to_datetime(employee_info["AL END DATE"])
                        pbar1.update(16)

                        pbar1.set_description("读取公共假期")
                        #ph_dates_df = df[3]
                        #parseGoogleHTMLSheet(ph_dates_df)
                        ph_dates_df = df[3].get_as_df()
                        ph_dates_df = ph_dates_df[ph_dates_df.iloc[:,0].astype(str).str.len() > 0]
                        ph_dates_df["PH DATE"] = pd.to_datetime(ph_dates_df["PH DATE"])
                        pbar1.update(16)

                        pbar1.set_description("读取手动录入假期")
                        #leaves_manual_df = df[4]
                        #parseGoogleHTMLSheet(leaves_manual_df)
                        leaves_manual_df = df[4].get_as_df()
                        leaves_manual_df = leaves_manual_df[leaves_manual_df.iloc[:,0].astype(str).str.len() > 0]
                        leaves_manual_df["DATE"] = pd.to_datetime(leaves_manual_df["DATE"])
                        leaves_manual_df["ID"] = leaves_manual_df["ID"].astype(int)
                        leaves_manual_df["ID"] = leaves_manual_df["ID"].astype(str)

                        pbar1.set_description("读取排班预览")
                        #previewSchedule = df[1]
                        #previewSchedule.drop("Unnamed: 0", axis=1, inplace=True)
                        #previewSchedule = previewSchedule.iloc[:20, :]

                        previewSchedule = df[1].get_as_df()

                        pbar1.set_description("完成")
                        pbar1.update(20)

                    print("排班主菜单")
                    startAction = option_num(["录入排班表数据库", "工资单时间分析", "移除或添加排班", "移除手动录入的假期记录", "生成排班表PDF", "退出排班"])
                    time.sleep(0.25)
                    userInputOne = option_limit(startAction, input("在这里输入>>>: "))

                    if userInputOne == 0:
                        with tqdm(total=100) as pbar2:
                            pbar2.set_description("开始")

                            OILPlusArr = np.zeros(len(shift_df))
                            PHPlusArr = np.zeros(len(shift_df))
                            ALPlusArr = np.zeros(len(shift_df))
                            CCLPlusArr = np.zeros(len(shift_df))

                            for index in range(len(shift_df)):
                                OILPlusArr[index] = assignLeaves(assign_for = "OIL",
                                                                employee_id=shift_df.iloc[index,0],
                                                                date=pd.to_datetime(shift_df.iloc[index, 1]),
                                                                employee_info_df=employee_info,
                                                                shift_df=shift_df,
                                                                leaves_manual_df=leaves_manual_df,
                                                                ph_dates_df = ph_dates_df
                                                                )

                                PHPlusArr[index] = assignLeaves(assign_for="PH",
                                                                employee_id = shift_df.iloc[index, 0],
                                                                date = pd.to_datetime(shift_df.iloc[index, 1]),
                                                                employee_info_df=employee_info,
                                                                shift_df = shift_df,
                                                                leaves_manual_df=leaves_manual_df,
                                                                ph_dates_df = ph_dates_df
                                                                )

                                ALPlusArr[index] = assignLeaves(assign_for="AL",
                                                                employee_id = shift_df.iloc[index, 0],
                                                                date = pd.to_datetime(shift_df.iloc[index, 1]),
                                                                employee_info_df=employee_info,
                                                                shift_df = shift_df,
                                                                leaves_manual_df=leaves_manual_df,
                                                                ph_dates_df = ph_dates_df
                                                                )

                                CCLPlusArr[index] = assignLeaves(assign_for="CCL",
                                                                employee_id = shift_df.iloc[index, 0],
                                                                date = pd.to_datetime(shift_df.iloc[index, 1]),
                                                                employee_info_df=employee_info,
                                                                shift_df = shift_df,
                                                                leaves_manual_df=leaves_manual_df,
                                                                ph_dates_df = ph_dates_df
                                                                )

                            shift_df["OIL PLUS"] = OILPlusArr
                            shift_df["PH PLUS"] = PHPlusArr
                            shift_df["AL PLUS"] = ALPlusArr
                            shift_df["CCL PLUS"] = CCLPlusArr

                            pbar2.set_description("准备完毕")
                            pbar2.update(16)

                            pbar2.set_description("移除干扰项")
                            dropExistingDateCriteria = shift_database["DATE"].isin(shift_df["DATE"].values)
                            dropIndex = shift_database[dropExistingDateCriteria].index
                            shift_database.drop(dropIndex, axis=0, inplace=True)
                            shift_database.reset_index(inplace=True)
                            shift_database.drop("index", axis=1, inplace=True)
                            pbar2.update(16)

                            pbar2.set_description("拼接旧记录")
                            for index in range(len(shift_df)):
                                if index > 0:
                                    shift_database = pd.concat([shift_database, DataFrame])

                                DataFrame = writeEntry(date=pd.to_datetime(shift_df.iloc[index, 1]),
                                                    employee_id=str(shift_df.iloc[index, 0]),
                                                    shift=shift_df.iloc[index, 2],
                                                    OILPlus=shift_df.iloc[index, 6],
                                                    PHPlus=shift_df.iloc[index, 7],
                                                    ALPlus=shift_df.iloc[index, 8],
                                                    CCLPlus=shift_df.iloc[index, 9],
                                                    remarks=shift_df.iloc[index, 5],
                                                    shift_database=shift_database,
                                                    employee_info_df=employee_info,
                                                    ph_dates_df=ph_dates_df
                                                    )
                            shift_database = pd.concat([shift_database, DataFrame])

                            pbar2.update(16)

                            pbar2.set_description("移除重复项")
                            shift_database.sort_values(by=["TIME LOG"], ascending=False, inplace=True, ignore_index=True)
                            dropIndex = shift_database[shift_database.duplicated(subset=["ID", "DATE"], keep="first")].index

                            if len(dropIndex) > 0:
                                shift_database.drop(dropIndex, axis=0, inplace=True)
                                shift_database.sort_values(by=["TIME LOG"], ascending=False, inplace=True)
                                shift_database.reset_index(inplace=True)
                                shift_database.drop("index", axis=1, inplace=True)

                            else:
                                pass

                            pbar2.update(16)

                            pbar2.set_description("写入云端")
                            shift_database["DATE"] = shift_database["DATE"].apply(lambda x: x.strftime("%Y-%m-%d"))
                            shift_database["IS PH"] = shift_database["IS PH"].apply(lambda x: str(x).strip().upper())
                            shift_database["TIME LOG"] = shift_database["TIME LOG"].apply(lambda x: x.strftime("%Y-%m-%d %H:%M:%S"))

                            NOW = pd.to_datetime(dt.datetime.now())
                            TWO_MONTHS_AGO = NOW - dt.timedelta(days=60)
                            TWO_MONTHS_AGO = TWO_MONTHS_AGO.strftime("%Y-%m-%d")
                            TWO_MONTHS_AGO = pd.to_datetime(TWO_MONTHS_AGO)

                            shift_db_upload = shift_database.copy()
                            shift_db_upload["DATE"] = pd.to_datetime(shift_db_upload["DATE"])
                            shift_db_upload["TIME LOG"] = pd.to_datetime(shift_db_upload["TIME LOG"])
                            shift_db_upload = shift_db_upload[shift_db_upload["DATE"] > TWO_MONTHS_AGO]

                            shift_db_upload["DATE"] = shift_db_upload["DATE"].apply(lambda y: y.strftime("%Y-%m-%d"))
                            shift_db_upload["TIME LOG"] = shift_db_upload["TIME LOG"].apply(lambda x: x.strftime("%Y-%m-%d %H:%M:%S"))

                            try:
                                shift_db = google_auth.open_by_url(database_url)
                                shift_db_sheetname_index = shift_db.worksheet(property="title", value=shift_db_sheetname).index
                                shift_db[shift_db_sheetname_index].set_dataframe(shift_db_upload, start="A1", nan="")
                                #print("为了节省云端资源和上传时间，")
                                #print("云端排班数据库仅保存最近两个月的记录")
                                #print("本地的数据库则会保存全时间全部数据")

                            except Exception as e:
                                print()
                                print("云端上传排班数据库失败，错误描述如下：")
                                print(e)
                                print()

                            pbar2.set_description("云端任务完成")
                            pbar2.update(16)

                            pbar2.set_description("写入本地")
                            drink_db = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_database_filename), sheet_name=drink_db_sheetname)
                            tabox_db = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_database_filename), sheet_name=takeaway_box_db_sheetname)
                            finance_db = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_database_filename), sheet_name=finance_db_sheetname)
                            promo_db = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_database_filename), sheet_name=promotion_db_sheetname)

                            tabox_stock_sheet = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_database_filename), sheet_name=tabox_stock_sheetname)
                            drink_stock_sheet = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_database_filename), sheet_name=drink_stock_sheetname)
                            rcv_sheet = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_database_filename), sheet_name=rcv_sheetname)

                            database_dfs = [drink_db, tabox_db, finance_db, promo_db, shift_database]

                            sheetnames = [drink_db_sheetname, takeaway_box_db_sheetname, finance_db_sheetname, promotion_db_sheetname,
                                        shift_db_sheetname]

                            with pd.ExcelWriter("{}/{}/{}".format(os.getcwd(), backup_foldername, local_database_filename)) as writer:
                                for index in range(len(database_dfs)):
                                    database_dfs[index]["DATE"] = pd.to_datetime(database_dfs[index]["DATE"])
                                    database_dfs[index]["TIME LOG"] = pd.to_datetime(database_dfs[index]["TIME LOG"])

                                    database_dfs[index]["DATE"] = database_dfs[index]["DATE"].apply(lambda y: y.strftime("%Y-%m-%d"))
                                    database_dfs[index]["TIME LOG"] = database_dfs[index]["TIME LOG"].apply(lambda x: x.strftime("%Y-%m-%d %H:%M:%S"))

                                    database_dfs[index].to_excel(writer, sheet_name=sheetnames[index], header=True, index=False)

                                tabox_stock_sheet.to_excel(writer, sheet_name=tabox_stock_sheetname, header=True, index=False)
                                drink_stock_sheet.to_excel(writer, sheet_name=drink_stock_sheetname, header=True, index=False)
                                rcv_sheet.to_excel(writer, sheet_name=rcv_sheetname, header=True, index=False)

                            pbar2.set_description("任务完成")
                            pbar2.update(20)

                        prtdf(shift_database.head(10))

                    elif userInputOne == 1:

                        shift_database = retrieve_shift_db(database_url, shift_db_sheetname, local_database_filename, backup_foldername)

                        if isinstance(shift_database, pd.DataFrame):
                            shift_database["DATE"] = pd.to_datetime(shift_database["DATE"])
                            shift_database["ID"] = shift_database["ID"].astype(int)
                            shift_database["ID"] = shift_database["ID"].astype(str)

                            payslip_send_channel = k_dict["payslip_time_send_channel"]

                            payslip_on_duty = str(payslip_on_duty).strip().upper()

                            rcv_telegram, rcv_email = get_rcv(fernet_key, k_dict, google_auth, backup_foldername, local_database_filename, False)

                            userInputTwo = 0
                            while userInputTwo != 3:
                                payslip_action = option_num(["整月分析", "自定义日期范围分析", "全部记录分析", "返回上一菜单"])
                                time.sleep(0.25)
                                userInputTwo = option_limit(payslip_action, input("在这里输入>>>: "))

                                if userInputTwo == 0:
                                    start_date, end_date = get_range_date(whole_month=True)

                                elif userInputTwo == 1:
                                    start_date, end_date = get_range_date(whole_month=False)

                                elif userInputTwo == 2:
                                    start_date = shift_database["DATE"].min()
                                    end_date = shift_database["DATE"].max()

                                if userInputTwo != 3:
                                    payslip_time_df = payslip_time(shift_database = shift_database,
                                                                start_date = pd.to_datetime(start_date),
                                                                end_date = pd.to_datetime(end_date),
                                                                employee_info_df=employee_info)

                                    payslip_time_df.reset_index(inplace=True)
                                    payslip_time_df.drop("index", axis=1, inplace=True)

                                    statement_dict = payslip_time_statement(concat_df = payslip_time_df,
                                                                            employee_info_df = employee_info,
                                                                            start_date = pd.to_datetime(start_date),
                                                                            end_date = pd.to_datetime(end_date) )


                                    prtdf(payslip_time_df)

                                    print()
                                    print("{}年{}月{}日至{}年{}月{}日{}员工工资工时分析单".format(start_date.year, start_date.month, start_date.day,
                                                                                            end_date.year, end_date.month, end_date.day,
                                                                                            shiftOutlet.strip().capitalize()))


                                    print()
                                    for k,i in statement_dict.items():
                                        print(k, ":", i)

                                    print()
                                    print()
                                    print("是否发送工时分析信息？")
                                    action_req = option_num(["是", "否"])
                                    time.sleep(0.25)
                                    userInputFive = option_limit(action_req, input("在这里输入>>>: "))
                                    if userInputFive == 0:
                                        send_string = "{}年{}月{}日至{}年{}月{}日{}员工工资工时分析单 \n ".format(start_date.year, start_date.month, start_date.day, end_date.year, end_date.month, end_date.day, shiftOutlet.strip().capitalize())
                                        send_string += "\n "
                                        for k, i in statement_dict.items():
                                            send_string += "{} : {} \n ".format(k, i)

                                        wifi = True

                                        if payslip_send_channel.strip().capitalize() == "Telegram":
                                            payslip_rcv = rcv_telegram[payslip_on_duty]
                                            telegram_api = fernet_decrypt(k_dict["payslip_time_telegram_bot_api"], fernet_key)
                                            sending_telegram(is_pr=False,
                                                             message=send_string,
                                                             api = telegram_api,
                                                             receiver=payslip_rcv,
                                                             wifi=wifi)

                                        elif payslip_send_channel.strip().capitalize() == "Email":
                                            email_server = fernet_decrypt(k_dict["payslip_time_email_server"], fernet_key)
                                            email_sender = fernet_decrypt(k_dict["payslip_time_email_sender"], fernet_key)
                                            email_sender_password = fernet_decrypt(k_dict["payslip_time_sender_password"], fernet_key)
                                            email_receiver = rcv_email[payslip_on_duty]

                                            sending_email(is_pr=False,
                                                          mail_server=email_server,
                                                          mail_sender=email_sender,
                                                          mail_sender_password=email_sender_password,
                                                          mail_receivers=email_receiver,
                                                          mail_subject="{}的工时分析".format(shiftOutlet),
                                                          message_string=send_string,
                                                          wifi=wifi)
                                        else:
                                            print("Payslip Time alert sending channel is not defined correctly")

                                    elif userInputFive == 1:
                                        pass



                        else:
                            print("无法获取排班数据库")

                    elif userInputOne == 2:
                        shift_ws = google_auth.open_by_url(shiftURL)
                        shift_sh_index = shift_ws.worksheet(property="title", value = shift_sheetname).index
                        shift_sh  = shift_ws[shift_sh_index]

                        shift_start_coordinate = "A2"
                        shift_end_coordinate = "C300"

                        shift_dr = pygsheets.datarange.DataRange(start=shift_start_coordinate, end=shift_end_coordinate, worksheet=shift_sh)

                        action_req = option_num(["移除", "添加"])
                        time.sleep(0.25)
                        userInputThree = option_limit(action_req, input("在这里输入>>>: "))

                        if userInputThree == 0:
                            action_req = option_num(["自定义日期范围移除", "全部移除"])
                            time.sleep(0.25)
                            userInputFour = option_limit(action_req, input("在这里输入>>>: "))

                            if userInputFour == 0:
                                current_shift_sh_values = shift_sh.get_values(start=shift_start_coordinate, end=shift_end_coordinate)
                                generated_df = pd.DataFrame(columns=["ID", "DATE", "SHIFT"], data=current_shift_sh_values)

                                if generated_df.empty:
                                    print("没有任何可移除项目")
                                elif len(generated_df) == 1:
                                    if (generated_df.iloc[0,0] == "") and (generated_df.iloc[0,1] == "") and (generated_df.iloc[0,2] == ""):
                                        print("没有任何可移除项目")
                                    else:
                                        print("排班请添加超过一项记录再使用此程序的移除功能")
                                else:
                                    generated_df["DATE"] = pd.to_datetime(generated_df["DATE"])
                                    generated_df["ID"] = generated_df["ID"].astype(int)
                                    generated_df["ID"] = generated_df["ID"].astype(str)

                                    start_date_remove, end_date_remove = get_range_date(whole_month=False)

                                    dateRange = pd.date_range(start=start_date_remove, end=end_date_remove)

                                    if generated_df[generated_df["DATE"].isin(dateRange)].empty:
                                        print("{}-{}之间没有任何项目".format(start_date_remove, end_date_remove))

                                    else:
                                        generated_df = generated_df[~generated_df["DATE"].isin(dateRange)]

                                        if generated_df.empty:
                                            shift_dr.clear()
                                            print("移除成功")

                                        else:
                                            generated_df.sort_values(by=["ID", "DATE"], ascending=True, ignore_index=True, inplace=True)
                                            generated_df["DATE"] = generated_df["DATE"].apply(lambda z : z.strftime("%Y-%m-%d"))

                                            uv = generated_df.values.tolist()

                                            shift_dr.clear()
                                            shift_dr.update_values(uv)
                                            print("移除成功")

                            elif userInputFour == 1:
                                shift_dr.clear()
                                print("全部移除成功")

                        elif userInputThree == 1:
                            start_date_add, end_date_add = get_range_date(whole_month=False)

                            dateRange = pd.date_range(start=start_date_add, end=end_date_add)

                            print("添加{}至{}的排班".format(start_date_add.strftime("%Y-%m-%d"), end_date_add.strftime("%Y-%m-%d")))
                            print("现在请选择要添加的员工ID, 请用逗号','间隔, 不能有多余的空格。")
                            print()
                            prtdf(employee_info.iloc[:, :2])
                            print()
                            take_input = input("在这里输入>>>: ")

                            em_ids = take_input.split(",")

                            recorded_em_id = employee_info["ID"].values.astype(str).tolist()

                            validate_id = True
                            for i in em_ids:
                                if i not in recorded_em_id:
                                    validate_id = False
                                else:
                                    pass

                            if validate_id:
                                dateRangeStr = []
                                for d in dateRange:
                                    dateRangeStr += [d.strftime("%Y-%m-%d")]

                                add_list = []
                                for i in em_ids:
                                    for d in dateRangeStr:
                                        add_list += [[str(i), str(d), ""]]

                            else:
                                print("ID错误! ")
                                add_list = []

                            if len(add_list) > 0:
                                if len(add_list) <= 298:
                                    current_shift_sh_values = shift_sh.get_values(start=shift_start_coordinate, end=shift_end_coordinate)
                                    generated_df = pd.DataFrame(columns=["ID", "DATE", "SHIFT"], data=current_shift_sh_values)

                                    if generated_df.empty:
                                        uv = add_list

                                    elif len(generated_df) == 1:
                                        if (generated_df.iloc[0,0] == "") and (generated_df.iloc[0,1] == "") and (generated_df.iloc[0,2] == ""):
                                            uv = add_list
                                        else:
                                            generated_df["DATE"] = pd.to_datetime(generated_df["DATE"])
                                            generated_df["ID"] = generated_df["ID"].astype(int)
                                            generated_df["ID"] = generated_df["ID"].astype(str)

                                            generated_df.sort_values(by=["ID", "DATE"], ascending=True, ignore_index=True, inplace=True)
                                            generated_df["DATE"] = generated_df["DATE"].apply(lambda z : z.strftime("%Y-%m-%d"))

                                            existing_values = generated_df.values.tolist()

                                            uv = existing_values + add_list
                                    else:
                                        generated_df["DATE"] = pd.to_datetime(generated_df["DATE"])
                                        generated_df["ID"] = generated_df["ID"].astype(int)
                                        generated_df["ID"] = generated_df["ID"].astype(str)

                                        generated_df.sort_values(by=["ID", "DATE"], ascending=True, ignore_index=True, inplace=True)
                                        generated_df["DATE"] = generated_df["DATE"].apply(lambda z : z.strftime("%Y-%m-%d"))

                                        existing_values = generated_df.values.tolist()

                                        uv = existing_values + add_list

                                    shift_dr.clear()
                                    shift_dr.update_values(uv)
                                    print("添加成功")
                                else:
                                    print("添加项目过多，最多加到第300行")
                            else:
                                pass

                    elif userInputOne == 3:
                        lm_start_coordinate = "A2"
                        lm_end_coordinate = "D60"

                        lm_ws = google_auth.open_by_url(shiftURL)
                        shift_leaves_manual_sheetname_index = lm_ws.worksheet(property="title", value=shift_leaves_manual_sheetname).index
                        lm_sh = lm_ws[shift_leaves_manual_sheetname_index]

                        dr = pygsheets.datarange.DataRange(start=lm_start_coordinate, end=lm_end_coordinate, worksheet=lm_sh)

                        dr.clear()
                        print("手动录入的假期记录已清除完成")

                    elif userInputOne == 4:
                        #isMonday = eval(str(previewSchedule.iloc[1,3]).capitalize())
                        isMonday = eval(str(previewSchedule.iloc[0,8]).capitalize())

                        if isMonday:
                            stock_count_foldername = "{}盘点文件".format(shiftOutlet.strip().capitalize())
                            songti_filename = "SongTi.ttf"
                            export_folderName = "{}预订导出".format(shiftOutlet.strip().capitalize())
                            logo_fileName = "SRX_logo.jpeg"

                            logo_path = "{}/{}/{}".format(os.getcwd(),export_folderName, logo_fileName)

                            if not os.path.exists("{}/{}/{}".format(os.getcwd(), stock_count_foldername, songti_filename)):
                                print("宋体TTF文件'{}'不存在, 排班表的生成无法继续".format(songti_filename))
                                print("请把宋体TTF文件'{}'保存至盘点文件名的目录下, 具体路径需在:'{}/{}/{}'".format(songti_filename, os.getcwd(), stock_count_foldername, songti_filename))

                            else:
                                if not os.path.exists(logo_path):
                                    print("公司标识图片'{}'不存在, 排班表的生成无法继续".format(logo_fileName))
                                    print("请把公司标识图片'{}'保存至预订导出的目录下, 具体路径需在: {}".format(logo_fileName, logo_path))
                                else:
                                    logoImagePath = pathlib.Path(logo_path)
                                    print("注意: 生成PDF之前先确保你已经录入最新的排班表数据,")
                                    print("否则假期天数将会出现偏差! ")
                                    schedule_pdf_options = option_num(["继续生成排班PDF", "返回上一菜单"])
                                    time.sleep(0.25)
                                    schedule_pdf_select = option_limit(schedule_pdf_options, input("在这里输入>>>: "))
                                    if schedule_pdf_select == 0:
                                        print("读取字体文件中...")
                                        custom_font_path = pathlib.Path("{}/{}/{}".format(os.getcwd(), stock_count_foldername, songti_filename))
                                        borb_custom_font = borb_TrueTypeFont.true_type_font_from_file(custom_font_path)
                                        print("字体文件读取完成。")
                                        print()
                                        generate_schedule_pdf(songTi=borb_custom_font, logoImagePath=logoImagePath, outlet=shiftOutlet, shift_database=shift_database, previewSchedule=previewSchedule, ph_dates_df=ph_dates_df)

                                    else:
                                        pass
                        else:
                            print("选择的日期不是星期一, 排班表无法继续生存。")

            else:
                print("无法获取排班数据库")

def schedule_font_color(text):
    if text in ["A", "B", "BG", "TS", "BD", "OT", "√", "√/", "/√"]:
        return borb_HexColor("#00308F")

    elif text in ["OFF", "PH", "OIL", "AL", "CCL", "NP", "MC"]:
        return borb_HexColor("#990F02")

    else:
        if text.find("/") != -1:
            findOFF = int(text.find("OFF"))
            findPH = int(text.find("PH"))
            findOIL = int(text.find("OIL"))
            findAL = int(text.find("AL"))
            findCCL = int(text.find("CCL"))
            findNP = int(text.find("NP"))
            findMC = int(text.find("MC"))

            findValueTotal = findOFF + findPH + findOIL + findAL + findCCL + findNP + findMC

            if findValueTotal == -7:
                return borb_HexColor("#00308F")

            else:
                return borb_HexColor("#979797")
        else:
            return borb_HexColor("#000000")

def generate_schedule_pdf(songTi, logoImagePath, outlet, shift_database, previewSchedule, ph_dates_df):
    with tqdm(total=100) as pbar:
        pbar.set_description("处理信息...")

        #monday = pd.to_datetime(str(previewSchedule.iloc[1,1]))
        monday = pd.to_datetime(str(previewSchedule.iloc[0,6]))

        #employee_id_show = previewSchedule.iloc[4:18, 0].astype(str).tolist()
        employee_id_show = previewSchedule.iloc[3:17,5].values.astype(str).tolist()

        #if "nan" in employee_id_show:
        #    employee_id_show.remove("nan")

        counter = 0
        for item in employee_id_show:
            if item == "":
                counter += 1
        
        if counter > 0:
            for t in range(counter):
                employee_id_show.remove("")
        else:
            pass

        if len(employee_id_show) < 14:
            for _ in range(14-len(employee_id_show)):
                employee_id_show += [" "]
        else:
            pass

        weekRange = [monday]
        for num in np.arange(1,7):
            weekRange += [ pd.to_datetime(monday + dt.timedelta(days=int(num)))]

        pbar.update(15)

        workRange = previewSchedule.copy()
        #workRange = workRange.iloc[4:18, 0:10]

        work_schedule_df = { #"序号" : np.arange(1, 15),
                            #"ID" : workRange.iloc[:, 0],
                            #"姓名" : workRange.iloc[:,1],
                            #"Mon" : workRange.iloc[:,2],
                            #"Tue" : workRange.iloc[:,3],
                            #"Wed" : workRange.iloc[:,4],
                            #"Thu" : workRange.iloc[:,5],
                            #"Fri" : workRange.iloc[:,6],
                            #"Sat" : workRange.iloc[:,7],
                            #"Sun" : workRange.iloc[:,8],
                            
                            "序号" : np.arange(1, 15),
                            "ID" : workRange.iloc[3:17, 5],
                            "姓名" : workRange.iloc[3:17,6],
                            "Mon" : workRange.iloc[3:17,7],
                            "Tue" : workRange.iloc[3:17,8],
                            "Wed" : workRange.iloc[3:17,9],
                            "Thu" : workRange.iloc[3:17,10],
                            "Fri" : workRange.iloc[3:17,11],
                            "Sat" : workRange.iloc[3:17,12],
                            "Sun" : workRange.iloc[3:17,13],
                            
                            }

        work_schedule_df["ID"] = work_schedule_df["ID"].astype(str)

        rest_days = {}

        keys = ["OIL", "PH", "AL", "CCL"]
        sunday = weekRange[-1].strftime("%Y-%m-%d")
        for key in keys:
            key_list = []
            for index in range(len(workRange.iloc[3:17, 5])):
                id = str(workRange.iloc[3:17, 5].values.astype(str)[index])
                id_string = "{}${}".format(id, sunday)
                if shift_database[shift_database["FOR VLOOKUP"] == id_string].empty:
                    key_list += [" "]
                else:
                    if float(str(shift_database[shift_database["FOR VLOOKUP"] == id_string][key].values[0])) == 0:
                        key_list += ["-"]
                    else:
                        key_list += [str(shift_database[shift_database["FOR VLOOKUP"] == id_string][key].values[0])]

            rest_days.update({ key : key_list})

        for key, value in rest_days.items():
            work_schedule_df.update( { key : value })

        work_schedule_df = pd.DataFrame(work_schedule_df)
        work_schedule_df["SIGN"] = np.repeat(" ", len(workRange.iloc[3:17,5]))
        work_schedule_df["RE"] = workRange.iloc[3:17,14].values
        work_schedule_df.drop("ID", axis=1, inplace=True)

        pbar.update(15)

        workerCountRange = previewSchedule.copy()
        #workerCountRange = workerCountRange.iloc[18:, :9]
        workerCountRange = workerCountRange.iloc[17:19,5:14]

        for num in range(6):
            workerCountRange["{}".format(num+1)] = np.repeat(" ", len(workerCountRange))

        pbar.update(15)

        #weekRemark = str(previewSchedule.iloc[18,9])
        weekRemark = str(previewSchedule.iloc[17,15])

        Document = borb_Document()
        Page = borb_Page(width=Decimal(842), height=Decimal(595))
        Document.add_page(Page)

        layout: borb_PageLayout = borb_SCL(Page)

        table0 = borb_Table(number_of_rows=1, number_of_columns=3)

        if int(weekRange[0].year) == int(weekRange[-1].year):
            table0.add(borb_Paragraph(str(weekRange[0].year) + "年", font=songTi, horizontal_alignment=borb_align.LEFT))
        else:
            table0.add(borb_Paragraph(str(weekRange[0].year)+"-"+str(weekRange[-1].year)+"年", font=songTi, horizontal_alignment=borb_align.LEFT))

        table0.add(borb_Paragraph("{}楼面排班表".format(outlet.strip().capitalize()), font=songTi, font_size=Decimal(22), horizontal_alignment=borb_align.CENTERED))
        table0.add(borb_Image(logoImagePath, width=Decimal(80), height=Decimal(20), horizontal_alignment=borb_align.RIGHT,))

        table0.set_padding_on_all_cells(Decimal(1), Decimal(1), Decimal(1), Decimal(1))
        table0.no_borders()
        layout.add(table0)
        pbar.update(15)

        table1 = borb_Table(number_of_rows=16, number_of_columns=15)

        table1_headers  = ["No",
                           "Name",
                           "Mon\n星期一",
                           "Tue\n星期二",
                           "Wed\n星期三",
                           "Thu\n星期四",
                           "Fri\n星期五",
                           "Sat\n星期六",
                           "Sun\n星期日",
                           "OIL",
                           "PH",
                           "AL",
                           "CCL",
                           "Sign",
                           "Remark"]

        for index in range(len(table1_headers)):
            if index in [2,3,4,5,6,7,8]:
                dateIndex = index - 2
                if weekRange[dateIndex] in ph_dates_df["PH DATE"].values:
                    table1.add(borb_Paragraph(table1_headers[index], font=songTi, horizontal_alignment=borb_align.CENTERED, font_color=borb_HexColor("#6F4E37"), background_color=borb_HexColor("#FDA4BA")))
                else:
                    table1.add(borb_Paragraph(table1_headers[index], font=songTi, horizontal_alignment=borb_align.CENTERED, font_color=borb_HexColor("#6F4E37")))
            else:
                table1.add(borb_Paragraph(table1_headers[index], font=songTi, horizontal_alignment=borb_align.CENTERED, font_color=borb_HexColor("#6F4E37")))

        table1.add(borb_Paragraph("序号", font=songTi, horizontal_alignment=borb_align.CENTERED, font_color=borb_HexColor("#6F4E37")))
        table1.add(borb_Paragraph("姓名", font=songTi, horizontal_alignment=borb_align.CENTERED, font_color=borb_HexColor("#6F4E37")))

        for day in weekRange:
            if day in ph_dates_df["PH DATE"].values:
                table1.add(borb_Paragraph(day.strftime("%m-%d"), font=songTi, horizontal_alignment=borb_align.CENTERED, background_color=borb_HexColor("#FDA4BA")))
            else:
                table1.add(borb_Paragraph(day.strftime("%m-%d"), font=songTi, horizontal_alignment=borb_align.CENTERED))

        for item in ["公休", "公期", "年假", "育儿假", "签名", "备注"]:
            table1.add(borb_Paragraph(item, font=songTi, horizontal_alignment=borb_align.CENTERED, font_color=borb_HexColor("#6F4E37")))

        for index in range(len(work_schedule_df)):
            for column in range(len(work_schedule_df.columns)):
                text = str(work_schedule_df.iloc[index, column])
                if text in ["nan", "", "None"]:
                    table1.add(borb_Paragraph(" ", font=songTi))
                else:
                    if column >= 9:
                        if column != 14:
                            try:
                                text = float(text)

                                if text == 0.5:
                                    text = "½"
                                else:
                                    if format(text, ".1f").endswith(".5"):
                                        text = "{}{}".format(int(text-0.5), "½")
                                    else:
                                        text = str(int(text))

                            except ValueError:
                                text = text

                            table1.add(borb_Paragraph(text, horizontal_alignment=borb_align.CENTERED, font_color=borb_HexColor("#00308F"), font_size=Decimal(15)))
                        else:
                            table1.add(borb_Paragraph(text, horizontal_alignment=borb_align.CENTERED, font_color=borb_HexColor("#00308F"), font=songTi))
                    else:
                        table1.add(borb_Paragraph(text, font=songTi, horizontal_alignment=borb_align.CENTERED, font_color=schedule_font_color(text)))

        table1.set_padding_on_all_cells(Decimal(2), Decimal(2), Decimal(2), Decimal(2))
        layout.add(table1)
        pbar.update(15)

        table2 = borb_Table(number_of_rows=2, number_of_columns=15)

        for index in range(len(workerCountRange)):
            for column in range(len(workerCountRange.columns)):
                text = str(workerCountRange.iloc[index, column])
                if text in ["员工", "数量", "早上", "晚上"]:
                    table2.add(borb_Paragraph(text, font=songTi, horizontal_alignment=borb_align.CENTERED, font_color=borb_HexColor("#6F4E37"), font_size=Decimal(15)))
                else:
                    table2.add(borb_Paragraph(text, font=songTi, horizontal_alignment=borb_align.CENTERED, font_color=borb_HexColor("#00308F"), font_size=Decimal(15)))

        table2.set_padding_on_all_cells(Decimal(1), Decimal(1), Decimal(1), Decimal(1))
        table2.no_borders()
        layout.add(table2)
        pbar.update(20)

        if len(weekRemark) > 0:
            if weekRemark not in ["nan", "-"]:
                layout.add(borb_Paragraph(weekRemark, font=songTi, horizontal_alignment=borb_align.RIGHT))
            else:
                pass
        else:
            pass

        schedule_export_folderName = "{}排班表PDF".format(outlet.strip().capitalize())
        schedule_pdf_filename = "{}楼面排班表{}月{}日至{}月{}日.pdf".format(outlet.strip().capitalize(), weekRange[0].month, weekRange[0].day, weekRange[-1].month, weekRange[-1].day)

        if not os.path.exists("{}/{}".format(os.getcwd(), schedule_export_folderName)):
            os.makedirs("{}/{}".format(os.getcwd(), schedule_export_folderName))

        if os.path.exists("{}/{}/{}".format(os.getcwd(), schedule_export_folderName, schedule_pdf_filename)):
            print("发现到你已经生成过'{}'PDF了, 是否覆盖? ".format(schedule_pdf_filename))
            overewrite_options = option_num(["覆盖(丢弃之前生成的PDF)", "不覆盖(保留之前生成的PDF)"])
            time.sleep(0.25)
            overewrite_select = option_limit(overewrite_options, input("在这里输入>>>: "))

            if overewrite_select == 0:
                with open("{}/{}/{}".format(os.getcwd(), schedule_export_folderName, schedule_pdf_filename), "wb") as pdf_file_handle:
                    borb_PDF.dumps(pdf_file_handle, Document)

                print("PDF任务完成。")

            else:
                print("保留之前生成的PDF。")
        else:
            with open("{}/{}/{}".format(os.getcwd(), schedule_export_folderName, schedule_pdf_filename), "wb") as pdf_file_handle:
                borb_PDF.dumps(pdf_file_handle, Document)

            print("PDF任务完成。")

        pbar.set_description("任务完成")
        pbar.update(5)

def coin_reset(google_auth, db_setting_url, constants_sheetname, serialized_rule_filename, backup_foldername):
    google_auth = google_auth
    db_setting_url = db_setting_url
    constants_sheetname = constants_sheetname
    serialized_rule_filename = serialized_rule_filename
    backup_foldername = backup_foldername

    fernet_key = get_key()

    if fernet_key == 0:
        print("安全密钥错误! ")
    else:
        outlet = get_outlet()
        k_dict = get_k_dictionary(google_auth, db_setting_url, constants_sheetname, serialized_rule_filename, outlet, fernet_key, backup_foldername)

        coin_url = fernet_decrypt(k_dict["coin_url"], fernet_key)

        ws = google_auth.open_by_url(coin_url)

        coin_values = ws[0].get_values(start="E3", end="F8")

        coins = {}

        for i in range(len(coin_values)):
            if float(coin_values[i][1]) == 0:
                pass
            else:
                coins.update({ str(coin_values[i][0]) : str(coin_values[i][1]) })

        clear_dr = pygsheets.datarange.DataRange(start="A2", end="C502", worksheet=ws[0])

        actions = option_num(["清理(保留各币种剩余数值)", "清除(清除所有记录)", "终止"])
        time.sleep(0.25)
        userInput = option_limit(actions, input("在这里输入>>>: "))

        if userInput == 0:
            TODAY = dt.datetime.today().strftime("%Y-%m-%d")

            uv_list = []
            for k,i in coins.items():
                uv_list += [[TODAY, i, k]]

            clear_dr.clear()
            clear_dr.update_values(uv_list)
            print("清理完成")

        elif userInput == 1:
            clear_dr.clear()
            print("清除完成")

        else:
            pass

def fernet_tool():
    fernet_key = get_key()

    if fernet_key == 0:
        print("安全密钥错误! ")

    else:
        f_handler = Fernet(fernet_key)

        actions = option_num(["加密", "解密", "终止"])
        time.sleep(0.25)
        userInput = option_limit(actions, input("在这里输入>>>: "))

        if userInput == 0:
            print("请粘贴要加密的文字, 按回车键执行")
            encrypt_string = input("在这里输入>>>: ")
            encrypt = f_handler.encrypt(encrypt_string.encode())

            print("加密文字如下: ")
            print(encrypt.decode())
            print()

        elif userInput == 1:
            print("请粘贴要解密的文字，按回车键执行")
            decrypt_string = input("在这里输入>>>: ")
            decrypt = fernet_decrypt(decrypt_string, fernet_key)

            print("解密文字如下: ")
            print(decrypt)
            print()

        else:
            pass

def email_validate(email_address):
    regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,7}\b'

    if (re.fullmatch(regex, email_address)):
        return True
    else:
        return False

def telegram_test_tool(google_auth, db_setting_url, constants_sheetname, serialized_rule_filename, backup_foldername):
    google_auth = google_auth
    db_setting_url = db_setting_url
    constants_sheetname = constants_sheetname
    serialized_rule_filename = serialized_rule_filename
    backup_foldername = backup_foldername

    fernet_key = get_key()

    if fernet_key == 0:
        print("安全密钥错误! ")
    else:
        outlet = get_outlet()
        k_dict = get_k_dictionary(google_auth, db_setting_url, constants_sheetname, serialized_rule_filename, outlet, fernet_key, backup_foldername)

        TIME_NOW = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        telegram_api = fernet_decrypt(k_dict["night_audit_telegram_bot_api"], fernet_key)

        print()
        print("输入你的测试信息，如果选择默认信息，直接按回车键")
        test_message = input("在这里输入>>>: ")

        if len(test_message.strip()) == 0:
            test_message = "这是机器人测试Telegram信息, 发送时间: {} \n This is Telegram Robot test message, sent at time: {}".format(TIME_NOW, TIME_NOW)
        else:
            pass

        print()
        print("输入你的Telegram Chat ID, 按回车键执行")
        chat_id_input = input("在这里输入>>>: ")

        sending_telegram(is_pr=False, message=test_message, api=telegram_api, receiver=chat_id_input, wifi=True)

def email_test_tool(google_auth, db_setting_url, constants_sheetname, serialized_rule_filename, backup_foldername):
    google_auth = google_auth
    db_setting_url = db_setting_url
    constants_sheetname = constants_sheetname
    serialized_rule_filename = serialized_rule_filename
    backup_foldername = backup_foldername

    fernet_key = get_key()

    if fernet_key == 0:
        print("安全密钥错误! ")
    else:
        outlet = get_outlet()
        k_dict = get_k_dictionary(google_auth, db_setting_url, constants_sheetname, serialized_rule_filename, outlet, fernet_key, backup_foldername)

        email_server = fernet_decrypt(k_dict["night_audit_email_server"], fernet_key)
        email_sender = fernet_decrypt(k_dict["night_audit_email_sender"], fernet_key)
        email_sender_password = fernet_decrypt(k_dict["night_audit_sender_password"], fernet_key)

        TIME_NOW = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        print("请输入电邮主题，如果选择默认主题，直接按回车键")
        email_subject = input("在这里输入>>>: ")

        if len(email_subject.strip()) == 0:
            email_subject = "电邮测试 {}".format(TIME_NOW)

        print("输入电邮内容, 如果选择默认内容, 直接按回车键")
        email_text = input("在这里输入>>>: ")

        if len(email_text.strip()) == 0:
            email_text = "电邮测试 {}".format(TIME_NOW)

        valid_email = False
        while not valid_email:
            email_receiver = input("输入收件人电子邮箱: ")

            if not email_validate(email_address=email_receiver):
                print("输入的'{}'不是电子邮箱，请重新输入。".format(email_receiver))

            else:
                valid_email = True

        sending_email(is_pr=False, mail_server=email_server, mail_sender=email_sender, mail_sender_password=email_sender_password, mail_receivers=email_receiver, mail_subject=email_subject, message_string=email_text, wifi=True)

def gen_drink_pdf(google_auth, db_setting_url, constants_sheetname, serialized_rule_filename, backup_foldername, database_url, custom_font, drink_num):
    google_auth = google_auth
    db_setting_url = db_setting_url
    constants_sheetname = constants_sheetname
    serialized_rule_filename = serialized_rule_filename
    backup_foldername = backup_foldername
    database_url = database_url
    custom_font = custom_font
    drink_num = drink_num

    fernet_key = get_key()

    if fernet_key == 0:
        print("安全密钥错误! ")
    else:
        outlet = get_outlet()
        k_dict = get_k_dictionary(google_auth, db_setting_url, constants_sheetname, serialized_rule_filename, outlet, fernet_key, backup_foldername)

        res = pyfiglet.figlet_format("Drink Report")
        print(res)
        time.sleep(0.15)

        drink_db_sheetname = k_dict["drink_database_sheetname"]
        local_db_filename = k_dict["local_database_filename"]

        database_url = fernet_decrypt(database_url, fernet_key)

        drink_pdf_foldername = "{}酒水明细PDF".format(outlet.strip().capitalize())

        if not os.path.exists("{}/{}".format(os.getcwd(), drink_pdf_foldername)):
            os.makedirs("{}/{}".format(os.getcwd(), drink_pdf_foldername))
            print("检测出未创建酒水明细表文件名, 已自动创建酒水明细表文件名")

        print("初始化...")
        try:
            ws = google_auth.open_by_url(database_url)
            drink_db_sheetname_index = ws.worksheet(property="title", value=drink_db_sheetname).index
            drink_db_online = ws[drink_db_sheetname_index].get_as_df()

        except Exception as e:
            print()
            print("通过机器人读取酒水数据失败，错误描述如下：")
            print(e)
            print()
            drink_db_online = None

        try:
            drink_db_local = pd.read_excel("{}/{}/{}".format(os.getcwd(), backup_foldername, local_db_filename), sheet_name=drink_db_sheetname)
        except Exception as e:
            print()
            print("本地读取酒水数据库失败，错误描述如下：")
            print(e)
            print()
            drink_db_local = None

        if isinstance(drink_db_online, pd.DataFrame) or isinstance(drink_db_local, pd.DataFrame):
            if isinstance(drink_db_online, pd.DataFrame):
                if drink_db_online.empty:
                    drink_db_online_max_date = None
                else:
                    drink_db_online["DATE"] = pd.to_datetime(drink_db_online["DATE"])
                    drink_db_online_max_date = drink_db_online["DATE"].max()

            else:
                drink_db_online_max_date = None

            if isinstance(drink_db_local, pd.DataFrame):
                if drink_db_local.empty:
                    drink_db_local_max_date = None
                else:
                    drink_db_local["DATE"] = pd.to_datetime(drink_db_local["DATE"])
                    drink_db_local_max_date = drink_db_local["DATE"].max()
            else:
                drink_db_local_max_date = None

            if (drink_db_online_max_date == None) and (drink_db_local_max_date != None):
                drink_database = drink_db_local.copy()

            elif (drink_db_local_max_date == None) and (drink_db_online_max_date != None):
                drink_database = drink_db_online.copy()

            elif (drink_db_online_max_date != None) and (drink_db_local_max_date != None):
                if drink_db_online_max_date >= drink_db_local_max_date:
                    drink_database = drink_db_online.copy()
                else:
                    drink_database = drink_db_local.copy()

            else:
                drink_database = None

        else:
            drink_database = None

        if isinstance(drink_database, pd.DataFrame):
            drink_database["DATE"] = pd.to_datetime(drink_database["DATE"])

            print("初始化完成")
            print()
            user_input = 0
            while user_input != 1:
                options = option_num(["整月生成", "退出酒水明细表PDF生成"])
                time.sleep(0.25)
                user_input = option_limit(options, input("在这里输入>>>: "))

                if user_input == 0:
                    start_date, end_date = get_range_date(whole_month=True)
                    date_range = pd.date_range(start=start_date, end=end_date)

                    drink_db_copy = drink_database.copy()
                    drink_db_copy["DATE"] = pd.to_datetime(drink_db_copy["DATE"])
                    drink_db_copy = drink_db_copy[drink_db_copy["DATE"].isin(date_range)]
                    drink_db_copy.sort_values(by=["DATE"], ascending=True, ignore_index=True, inplace=True)

                    if drink_db_copy.empty:
                        print("{}至{}没有任何酒水记录, 酒水明细PDF无法生存".format(start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d")))

                    else:
                        if len(drink_db_copy) != len(date_range):
                            print("{}年{}月的酒水明细没有整月完整的信息, 是否继续生成? ".format(start_date.year, start_date.month))
                            select_options = option_num(["是", "否"])
                            time.sleep(0.25)
                            select_input = option_limit(select_options, input("在这里输入>>>: "))

                            if select_input == 0:
                                continue_gen = True
                            else:
                                continue_gen = False
                        else:
                            continue_gen = True

                        drink_pdf_filename = "{}年{}月{}酒水明细表.pdf".format(start_date.year, start_date.month, outlet.strip().capitalize())

                        if os.path.exists("{}/{}/{}".format(os.getcwd(), drink_pdf_foldername, drink_pdf_filename)):
                            print("{}年{}月{}的酒水明细表已存在，是否生成新的来覆盖?".format(start_date.year, start_date.month, outlet.strip().capitalize()))
                            print("如果你选择了覆盖, 之前的明细表会被系统删除, 此操作不可逆转")
                            print()
                            options = option_num(["继续生成明细表(覆盖旧文件)", "终止生成(保留旧文件)"])
                            time.sleep(0.25)
                            user_input = option_limit(options, input("在这里输入>>>: "))

                            if user_input == 0:
                                pass
                            else:
                                continue_gen = False
                        else:
                            pass

                        if continue_gen:
                            print("生成酒水明细表PDF过程比较漫长, 请耐心等待...")
                            with tqdm(total=100) as pbar:
                                pbar.set_description("获取酒水名称...")

                                drink_names = []
                                for n in range(drink_num):
                                    if str(k_dict["drink{}".format(n)]).startswith("饮料"):
                                        continue
                                    else:
                                        drink_names += [str(k_dict["drink{}".format(n)])]

                                table_columns = ["日期", "上日存货", "进", "出", "实结存", "签名", "备注"]

                                pbar.update(20)

                                pbar.set_description("开始生成PDF...")
                                Document = borb_Document()
                                Page = borb_Page(width=Decimal(595), height=Decimal(842))
                                Document.add_page(Page)

                                layout: borb_PageLayout = borb_MCL(Page)

                                drink_db_copy["DATE"] = drink_db_copy["DATE"].apply(lambda z : z.strftime("%d"))

                                tables_dict = {}
                                table_counter = 0
                                df_dict = {}
                                title_dict = {}
                                for index in range(len(drink_names)):
                                    dataframe = pd.DataFrame({
                                        table_columns[0]: drink_db_copy["DATE"].values,
                                        table_columns[1]: drink_db_copy["{}上日存货".format(drink_names[index])].values,
                                        table_columns[2]: drink_db_copy["{}进".format(drink_names[index])].values,
                                        table_columns[3]: drink_db_copy["{}出".format(drink_names[index])].values,
                                        table_columns[4]: drink_db_copy["{}实结存".format(drink_names[index])].values,
                                        table_columns[5]: np.repeat(np.nan, len(drink_db_copy["DATE"].values)),
                                        table_columns[6]: drink_db_copy["{}备注".format(drink_names[index])].values,

                                    })

                                    df_dict.update({ "table{}".format(table_counter) : dataframe})

                                    tables_dict.update({ "table{}".format(table_counter) : borb_flex_Table(number_of_rows=32, number_of_columns=len(table_columns)) })

                                    title_dict.update({ "table{}".format(table_counter) : "{}年{}月份{}明细".format(start_date.year, start_date.month, drink_names[index])})

                                    table_counter += 1

                                pbar.update(20)

                                #table manipulation
                                for key in tables_dict:
                                    space_avail = 32*len(table_columns)

                                    for c in table_columns:
                                        tables_dict[key].add(borb_TableCell(
                                                                borb_Paragraph(c, font=custom_font, horizontal_alignment=borb_align.CENTERED)))

                                    space_avail -= len(table_columns)

                                    for row in range(len(df_dict[key])):
                                        for column in range(len(df_dict[key].columns)):
                                            if str(df_dict[key].iloc[row, column]) == "nan":
                                                tables_dict[key].add(borb_Paragraph(" ", font=custom_font, horizontal_alignment=borb_align.CENTERED))
                                            else:
                                                tables_dict[key].add(borb_Paragraph(str(df_dict[key].iloc[row, column]), font=custom_font, horizontal_alignment=borb_align.CENTERED))

                                            space_avail -= 1

                                    if space_avail > 0:
                                        for extraSpace in range(space_avail):
                                            tables_dict[key].add(borb_Paragraph(" ", font=custom_font, horizontal_alignment=borb_align.CENTERED))
                                    else:
                                        pass

                                    tables_dict[key].set_padding_on_all_cells(Decimal(1), Decimal(1), Decimal(1), Decimal(1))

                                pbar.update(20)

                                #append to layout
                                for key in tables_dict:
                                    layout.add(borb_Paragraph("三人行酒水明细表({})".format(outlet.strip().capitalize()), font=custom_font, horizontal_alignment=borb_align.CENTERED))

                                    layout.add(borb_Paragraph(str(title_dict[key]), font=custom_font, horizontal_alignment=borb_align.CENTERED))

                                    layout.add(tables_dict[key])

                                pbar.update(35)
                                pbar.set_description("PDF保存中...")

                                drink_pdf_filename = "{}年{}月{}酒水明细表.pdf".format(start_date.year, start_date.month, outlet.strip().capitalize())

                                with open("{}/{}/{}".format(os.getcwd(), drink_pdf_foldername, drink_pdf_filename), "wb") as pdf_file_handle:
                                    borb_PDF.dumps(pdf_file_handle, Document)

                                pbar.set_description("PDF任务完成")
                                pbar.update(5)

                        else:
                            pass

                else:
                    pass

            else:
                pass

        else:
            print("无法获取任何酒水数据库, 酒水明细表PDF生成无法继续")

def trigger_gen_drink_pdf(k_dict, date_dict):
    dfb = date_dict["dfb"]

    LAST_DAY = month_last_day(dfb.year, dfb.month)
    MONTH_LAST_DAY = pd.to_datetime(dt.datetime(dfb.year, dfb.month, LAST_DAY))

    is_auto = eval(str(k_dict["auto_gen_drink_pdf"]).strip().capitalize())

    if is_auto:
        if dfb == MONTH_LAST_DAY:
            return True
        else:
            return False
    else:
        return False

def float_check(x):
    try:
        x = float(x)
        return True
    except ValueError:
        return False

def rtn_whatsapp_sender(time_slots, google_auth):
    from selenium import webdriver
    from sys import platform as sel_platform

    print("Please provide reservation dataframe URL")
    print("请提供预订表格的网址")
    time.sleep(0.15)
    url = input("在这里输入>>>: ")

    print()
    print()
    print("Please provide outlet")
    print("请提供门店")
    time.sleep(0.15)
    outlet = input("在这里输入>>>: ")

    print()
    print("您输入的URL是: {}".format(url))
    print()
    print("您输入的门店是: {}".format(outlet))
    print()
    input("按回车键继续>>>: ")
    print()
    print("Fetching Data...Please Wait...")
    print("提取数据中...请稍等...")
    print()

    df = google_auth.open_by_url(url)[0]
    df = df.get_as_df()

    df["姓名"] = df["姓名"].astype(str)
    df["电话"] = df["电话"].astype(str)
    df["轮数"] = df["轮数"].astype(str)
    df["桌号"] = df["桌号"].astype(str)
    df["是否完成订单"] = df["是否完成订单"].astype(str)

    df = df[df["是否完成订单"] != "TRUE"]
    df = df[df["属性"] == "堂食"]
    df = df[df["轮数"] != "无效"]

    df["时间"] = df["轮数"].apply(lambda x: time_slots[x])

    df["WhatsApp"] = "尊敬的"+df["姓名"]+"您好! 这里是三人行("+str(outlet).strip().upper()+")中餐馆,温馨通知您在除夕的订位是"+df["轮数"]+"("+df["时间"]+"),您的桌号是"+df["桌号"]+"。谢谢! 祝您新年快乐! Greetings, This is San Ren Xing("+str(outlet).strip().upper()+"), I'd like to inform you that your time slot for CNY eve is " + df["时间"] + ". Your table number is/are " + df["桌号"] + ". Thank you, wishing you a happy Chinese New Year! "

    print()
    print()
    prtdf(df)
    print()
    if len(df) <= 0:
        pass
    else:
        print("样本: ")
        for line in range(len(df.iloc[-1, :])):
            print("{}: {}".format(df.columns[line], df.iloc[-1,:][line]))

    print()
    print()
    input("按回车键继续>>>: ")

    phone = df["电话"].values.astype(str)
    texts = df["WhatsApp"].values.astype(str)

    total_numbers = len(phone)
    print("##########################################################")
    print('共找到{}个电话号码'.format(total_numbers))
    print("##########################################################")
    print()
    print()
    print()
    if total_numbers > 0:
        print('确保你的电脑安装了对应的浏览器和有网络连接')
        print('WhatsApp开启后, 请登录你的账号。')
        print()
        print()
    
        print("Please select the browser you'd like to use")
        print("请选择您偏爱使用的浏览器")
        print()
        time.sleep(0.15)
        browser_options = option_num(["谷歌浏览器Chrome", "火狐浏览器Firefox", "微软浏览器Edge"])
        time.sleep(0.25)
        browser_select = option_limit(browser_options, input("在这里输入>>>: "))
    
        if browser_select == 0:
            print("If you don't have chromedriver installed or had experienced outdated chromedriver error, ")
            print("please download it at: https://developer.chrome.com/docs/chromedriver/downloads")
            print("如果您没有安装chromedriver驱动或者驱动版本过老, 请访问上述链接来下载该驱动。")
            print()
            input("按回车键继续>>>: ")
            print()
            print("Please provide chromedriver's absolute path on your computer: ")
            print("请提供chromedriver驱动的绝对路径: ")
            chromedriver_path = input("在这里输入>>>: ")
            
            from selenium import webdriver
            from selenium.webdriver.chrome.service import Service
            from selenium.webdriver.chrome.options import Options
    
            # Set up Chrome options
            chrome_options = Options()
            
            # Optional: start Chrome in incognito mode
            chrome_options.add_argument("--incognito")
            
            # Path to chromedriver (adjust for your system)
            service = Service(chromedriver_path.strip())  
            
            # Initialize Chrome driver
            driver = webdriver.Chrome(service=service, options=chrome_options)
                
        elif browser_select == 1:
            from selenium.webdriver.firefox.service import Service as FirefoxService
            from webdriver_manager.firefox import GeckoDriverManager
    
            service = FirefoxService(GeckoDriverManager().install())
            options = webdriver.FirefoxOptions()
    
        
            options.add_argument("-private")
            driver = webdriver.Firefox(service=service, options=options)
    
        elif browser_select == 2:
            print("If you don't have msedgedriver installed,")
            print("please download it at: https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver?form=MA13LH")
            print("如果您没有安装msedgedriver驱动, 请访问上述链接来下载该驱动。")
            print()
            input("按回车键继续>>>: ")
            print()
            print("Please provide msedgedriver's absolute path on your computer: ")
            print("请提供msedgedriver驱动的绝对路径: ")
            msedgedriver_path = input("在这里输入>>>: ")
            
            from selenium import webdriver
            from selenium.webdriver.edge.service import Service
            from selenium.webdriver.edge.options import Options
            
            # Set up Edge options
            edge_options = Options()
            edge_options.use_chromium = True  # Ensure using Chromium-based Edge
            
            # Path to your msedgedriver executable
            service = Service(msedgedriver_path.strip())
            
            # Initialize the Edge driver
            driver = webdriver.Edge(service=service, options=edge_options)
            
        driver.get('https://web.whatsapp.com')
        print("当你看到你的聊天列表的时候再按回车键继续。")
        input("按回车键继续>>>: ")
    
        for index in range(len(phone)):
            if str(phone[index]).startswith("+"):
                isLocalPhone = False
            else:
                isLocalPhone = True
    
            number = str(phone[index])
            text = str(texts[index])
    
            if isLocalPhone:
                url = "https://web.whatsapp.com/send?phone=65" + number + "&text=" + text
            else:
                url = "https://web.whatsapp.com/send?phone=" + number.replace("+", "") + "&text" + text
    
            driver.get(url)
    
            whatsapp_options = option_num(["发送成功, 发送下一个", "发送失败, 重试一次", "发送失败, 跳过, 发送下一个"])
            time.sleep(0.25)
            whatsapp_select = option_limit(whatsapp_options, input("在这里输入>>>: "))
    
            if whatsapp_select == 0:
                print("{}发送成功! ".format(number))
                continue
    
            elif whatsapp_select == 1:
                print("{}发送失败, 将重试一次。".format(number))
                for _ in range(1):
                    print("重新发给{}。".format(number))
                    if isLocalPhone:
                        url = "https://web.whatsapp.com/send?phone=65" + number + "&text=" + text
                    else:
                        url = "https://web.whatsapp.com/send?phone=" + number.replace("+", "") + "&text" + text
    
                    driver.get(url)
                    input("按回车键继续>>>: ")
                    continue
    
            elif whatsapp_select == 2:
                print("{}发送失败, 已跳过。".format(number))
                continue
    
        print("任务完成")
        driver.quit()
    else:
        pass
    
    return df

def inline_rsv_whatsapp_sender():
    from selenium import webdriver as sel_webdriver
    from selenium.webdriver.chrome.options import Options as sel_Options
    from webdriver_manager.chrome import ChromeDriverManager
    from sys import platform as sel_platform
    from selenium.webdriver.chrome.service import Service as sel_service

    print("请选择Inline订位文件")
    print()
    print()
    fileShown = os.listdir()
    fileToShow = []

    for file in fileShown:
        if file.startswith("~"):
            pass
        elif file.startswith("."):
            pass
        else:
            fileToShow += [file]

    fileToShow += ["返回上一菜单"]

    selectFileToViewOptions = option_num(fileToShow)
    time.sleep(0.25)
    userInputEight = option_limit(selectFileToViewOptions, input("在这里输入>>>: "))

    if userInputEight == len(fileToShow)-1:
        return 2
    else:
        fileName = fileToShow[userInputEight]
        df = pd.read_excel(fileName)

        drop_index = []
        for index in range(len(df)):
            text = str(df.iloc[index, 13])
            status = str(df.iloc[index, 16])

            if len(text) == 0:
                pass
            else:
                for char in text:
                    if char == "甭":
                        drop_index += [index]

            if "Canceled" in status:
                drop_index += [index]

        if len(drop_index) > 0:
            df.drop(drop_index, axis=0, inplace=True)

        df["Adults"] = df["Adults"].astype(int)
        df["Kids"] = df["Kids"].astype(int)
        df["Highchairs"] = df["Highchairs"].astype(int)
        df["TOTAL PAX"] = df["Adults"] + df["Kids"] + df["Highchairs"]
        df["RSV Time"] = pd.to_datetime(df["RSV Time"], format="%m/%d/%Y, %I:%M %p")
        df["RSV Time"] = df["RSV Time"].dt.strftime("%d %b %H:%M")
        df["TOTAL PAX"] = df["TOTAL PAX"].astype(str)
        df["Patron Phone No."] = df["Patron Phone No."].astype(str)

        df["message"] = "Hi "+df["Patron Name"]+", we noticed that you have a reservation with us on "+df["RSV Time"]+" for "+df["TOTAL PAX"]+", is it confirmed that y'all are coming? 😁"
        prtdf(df)
        print()
        if len(df) <= 0:
            pass
        else:
            print("样本: ")
            for line in range(len(df.iloc[-1,:])):
                print("{}: {}".format(df.columns[line], df.iloc[-1,:][line]))
            print()
            print()
        input("按回车键继续>>>: ")

        options = sel_Options()

        if sel_platform == "win32":
            options.binary_location = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

        total_numbers = len(df)
        print("##########################################################")
        print('共找到{}个电话号码'.format(total_numbers))
        print("##########################################################")
        print()
        print()
        print()
        print('确保你的电脑安装了谷歌浏览器和有网络连接')
        service = sel_service()
        options = sel_webdriver.ChromeOptions()
        driver = sel_webdriver.Chrome(service=service, options=options)
        #driver = sel_webdriver.Chrome(ChromeDriverManager().install())
        print('WhatsApp开启后, 请登录你的账号。')
        driver.get('https://web.whatsapp.com')
        print("当你看到你的聊天列表的时候再按回车键继续。")
        input("按回车键继续>>>: ")

        for index in range(len(df)):
            number = str(df["Patron Phone No."].values[index])
            text = str(df["message"].values[index])

            url = "https://web.whatsapp.com/send?phone=" + number + "&text=" + text

            driver.get(url)

            whatsapp_options = option_num(["发送成功, 发送下一个", "发送失败, 重试一次", "发送失败, 跳过, 发送下一个"])
            time.sleep(0.25)
            whatsapp_select = option_limit(whatsapp_options, input("在这里输入>>>: "))

            if whatsapp_select == 0:
                print("{}发送成功! ".format(number))
                continue

            elif whatsapp_select == 1:
                print("{}发送失败, 将重试一次。".format(number))
                for _ in range(1):
                    print("重新发给{}。".format(number))
                    url = "https://web.whatsapp.com/send?phone=" + number + "&text=" + text

                    driver.get(url)

                    input("按回车键继续>>>: ")
                    continue

            elif whatsapp_select == 2:
                print("{}发送失败, 已跳过。".format(number))
                continue

        print("任务完成")
        driver.quit()

def uniform_main(google_auth, db_setting_url, constants_sheetname, serialized_rule_filename, backup_foldername):

    res = pyfiglet.figlet_format("Uniform")
    print(res)
    time.sleep(0.15)

    on_net = on_internet()

    if not on_net:
        print()
        print()
        print("无网络连接。")
        print("制服库存的查看需要全程连接网络。")

    else:
        fernet_key = get_key()

        if fernet_key == 0:
            print("安全密钥错误! ")
        else:
            outlet = get_outlet()
            k_dict = get_k_dictionary(google_auth, db_setting_url, constants_sheetname, serialized_rule_filename, outlet, fernet_key, backup_foldername)
            uniformURL = fernet_decrypt(k_dict["uniform_sheet_url"], fernet_key)

            df = google_auth.open_by_url(uniformURL)

            uniform_df = df[1].get_as_df()
            uniform_df = uniform_df[uniform_df.iloc[:, 0].astype(str).str.len() > 0]

            print()
            print("报告生成时间: {}".format(dt.datetime.now().strftime("%Y年%m月%d日 %H:%M:%S")))
            print("{}各尺码库存详情: ".format(outlet.strip().capitalize()))
            print("尺码: 库存量")
            for index in range(len(uniform_df)):
                print("{}: {}".format(uniform_df.iloc[index, 0], uniform_df.iloc[index, 1]))
            print()
            print()
            prtdf(uniform_df)
            print()
            print()

def main(database_url, db_setting_url, serialized_rule_filename, service_filename, constants_sheetname, google_auth, box_num, drink_num, promo_num, lun_sales, lun_gc, tb_sales, tb_gc, lun_fwc, lun_kwc, tb_fwc, tb_kwc, night_fwc, night_kwc, script_backup_filename, script, wifi, backup_foldername,cashier_on_duty, drink_on_duty, box_on_duty, payslip_on_duty, do_not_show_menu):
    database_url = database_url
    db_setting_url = db_setting_url
    serialized_rule_filename = serialized_rule_filename
    service_filename = service_filename
    constants_sheetname = constants_sheetname
    google_auth = google_auth
    box_num = box_num
    drink_num = drink_num
    promo_num = promo_num
    lun_sales = lun_sales
    lun_gc = lun_gc
    tb_sales = tb_sales
    tb_gc = tb_gc
    lun_fwc = lun_fwc
    lun_kwc = lun_kwc
    tb_fwc = tb_fwc
    tb_kwc = tb_kwc
    night_fwc = night_fwc
    night_kwc = night_kwc
    script_backup_filename = script_backup_filename
    script = script
    wifi = wifi
    backup_foldername = backup_foldername
    cashier_on_duty = cashier_on_duty
    drink_on_duty = drink_on_duty
    box_on_duty = box_on_duty
    payslip_on_duty = payslip_on_duty
    do_not_show_menu = do_not_show_menu

    if do_not_show_menu:
        night_audit_main(database_url, db_setting_url, serialized_rule_filename, service_filename, constants_sheetname, google_auth, box_num, drink_num, promo_num, lun_sales, lun_gc, tb_sales, tb_gc, lun_fwc, lun_kwc, tb_fwc, tb_kwc, night_fwc, night_kwc, script_backup_filename, script, wifi, backup_foldername,cashier_on_duty, drink_on_duty, box_on_duty, payslip_on_duty, manager_on_duty)

    else:
        if int(np.random.randint(1,7,size=1)[0]) % 2 == 0:
            backup_script(script_backup_filename, script, backup_foldername)
        else:
            pass

        res = pyfiglet.figlet_format("San Ren Xing Super App")
        print(res)
        time.sleep(0.15)

        SRX_take_input = 0
        while SRX_take_input != 7:
            print()
            print("Main Menu")
            options = option_num(["关帐", "盘点", "排班", "生成酒水明细表", "WhatsApp", "制服库存", "工具箱", "终止Super App"])
            time.sleep(0.25)
            SRX_take_input = option_limit(options, input("在这里输入>>>: "))

            if SRX_take_input == 0:
                night_audit_main(database_url, db_setting_url, serialized_rule_filename, service_filename, constants_sheetname, google_auth, box_num, drink_num, promo_num, lun_sales, lun_gc, tb_sales, tb_gc, lun_fwc, lun_kwc, tb_fwc, tb_kwc, night_fwc, night_kwc, script_backup_filename, script, wifi, backup_foldername,cashier_on_duty, drink_on_duty, box_on_duty, payslip_on_duty, manager_on_duty)

            elif SRX_take_input == 1:
                inventory_main(google_auth, db_setting_url, constants_sheetname, serialized_rule_filename, backup_foldername, database_url, box_num, drink_num)

            elif SRX_take_input == 2:
                work_schedule_main(google_auth, db_setting_url, constants_sheetname, serialized_rule_filename, backup_foldername, database_url, payslip_on_duty)

            elif SRX_take_input == 3:
                outlet = get_outlet()
                stock_count_foldername = "{}盘点文件".format(outlet.strip().capitalize())
                songti_filename = "SongTi.ttf"

                if not os.path.exists("{}/{}/{}".format(os.getcwd(), stock_count_foldername, songti_filename)):
                    print("宋体TTF文件'{}'不存在, 酒水明细表的生成无法继续".format(songti_filename))
                    print("请把宋体TTF文件'{}'保存至盘点文件名的目录下, 具体路径需在:'{}/{}/{}'".format(songti_filename, os.getcwd(), stock_count_foldername, songti_filename))

                else:
                    print("读取字体文件中...")
                    custom_font_path = pathlib.Path("{}/{}/{}".format(os.getcwd(), stock_count_foldername, songti_filename))
                    borb_custom_font = borb_TrueTypeFont.true_type_font_from_file(custom_font_path)
                    print("字体文件读取完成。")
                    print()
                    gen_drink_pdf(google_auth, db_setting_url, constants_sheetname, serialized_rule_filename, backup_foldername, database_url, borb_custom_font, drink_num)

            elif SRX_take_input == 4:
                res = pyfiglet.figlet_format("WhatsApp")
                print(res)
                time.sleep(0.15)
                send_whatsapp_options = option_num(["发预订WhatsApp", "发确认定位WhatsApp", "返回上一菜单"])
                time.sleep(0.25)
                wa_take_input = option_limit(send_whatsapp_options, input("在这里输入>>>: "))

                while wa_take_input != 2:
                    if wa_take_input == 0:
                        #edit the time slots over here
                        time_slots = {"第1轮": "17:00-18:20",
                                      "第2轮": "18:30-20:00",
                                      "第3轮": "20:10-22:00"}

                        rtn_whatsapp_sender(time_slots, google_auth)

                    elif wa_take_input == 1:
                        wa_take_input = inline_rsv_whatsapp_sender()
                else:
                    pass

            elif SRX_take_input == 5:
                uniform_main(google_auth, db_setting_url, constants_sheetname, serialized_rule_filename, backup_foldername)

            elif SRX_take_input == 6:

                res = pyfiglet.figlet_format("Tool Box")
                print(res)
                time.sleep(0.15)

                on_net = on_internet()

                if not on_net:
                    print("没有网络连接")
                    print("工具箱需全程连接网络来使用")

                else:
                    toolbox_input = 0
                    while toolbox_input != 3:
                        toolboxes = option_num(["硬币Sheet清理工具", "Fernet加密解密工具", "测试发送信息", "关闭工具箱"])
                        time.sleep(0.25)
                        toolbox_input = option_limit(toolboxes, input("在这里输入>>>: "))

                        if toolbox_input == 0:
                            coin_reset(google_auth, db_setting_url, constants_sheetname, serialized_rule_filename, backup_foldername)

                        elif toolbox_input == 1:
                            fernet_tool()

                        elif toolbox_input == 2:
                            send_test_input = 0
                            while send_test_input != 2:
                                print()
                                print()
                                options = option_num(["发送Telegram测试", "发送电邮测试", "退出发送测试"])
                                time.sleep(0.25)
                                send_test_input = option_limit(options, input("在这里输入>>>: "))

                                if send_test_input == 0:
                                    telegram_test_tool(google_auth, db_setting_url, constants_sheetname, serialized_rule_filename, backup_foldername)

                                elif send_test_input == 1:
                                    email_test_tool(google_auth, db_setting_url, constants_sheetname, serialized_rule_filename, backup_foldername)

        else:
            res = pyfiglet.figlet_format("Thank You")
            print(res)

if __name__ == "__main__":
    main(database_url, db_setting_url, serialized_rule_filename, service_filename, constants_sheetname, google_auth, box_num, drink_num, promo_num, lun_sales, lun_gc, tb_sales, tb_gc, lun_fwc, lun_kwc, tb_fwc, tb_kwc, night_fwc, night_kwc, script_backup_filename, script, wifi, backup_foldername,cashier_on_duty, drink_on_duty, box_on_duty, payslip_on_duty, do_not_show_menu)
