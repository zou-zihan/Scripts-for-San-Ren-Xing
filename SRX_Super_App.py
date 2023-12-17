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

#打包盒管理负责人(加引号)
box_on_duty = ""

#酒水管理负责人(加引号)
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

#下午茶营业额
tb_sales = 0.0

#下午茶顾客人数
tb_gc = 0

#下午茶楼面员工人数(加引号)
tb_fwc = ""

#下午茶厨房员工人数(加引号)
tb_kwc = ""

#晚上楼面员工人数(加引号)
night_fwc = ""

#晚上厨房员工人数(加引号)
night_kwc = ""

#收银员(加引号)
cashier_on_duty = ""

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
    options_list = np.array(options_list)
    rx = np.arange(len(options_list)).astype(str)
    kou = np.repeat('扣', len(options_list))
    options_list_kou = np.core.defchararray.add(options_list,kou)
    options_list_kou_int = np.core.defchararray.add(options_list_kou, rx)
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

            try:
                tsbt_index = int(read[read["0"] == 'Total Sales Before Tax & Srv Chg'].index[0])

            except IndexError:
                tsbt_index = int(read[read["0"] == "Total Sales Before Tax"].index[0])

            read.iloc[tsbt_index, 0] = 'Total Sales Before Tax & Srv Charge'

            gst_index = int(read[read["0"] == "GST 8%"].index[0])
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
            try:
                tabox_url = fernet_decrypt(str(k_dict["box_drink_in_out_url"]), fernet_key).strip()
                tabox_url += "htmlview"

                tabox_df = pd.read_html(tabox_url, encoding="utf-8")[0]
                parseGoogleHTMLSheet(tabox_df)

                print("已通过仅读模式在网络获取打包盒出入库数据。")
                print("注意在没有机器人的帮助下，所有的打包盒名字和出入库的数据不能自动重置。")

            except Exception as e:
                print()
                print("通过仅读模式在网络获取打包盒出入库数据失败, 错误描述如下: ")
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

def parse_print_rule(value_dict, rule_df_dict, write_finance_db, write_promo_db, write_drink_db, tabox_write_db):

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

    print_error = False
    for booleans in [write_finance_db,write_promo_db,write_drink_db,tabox_write_db]:
        if not booleans:
            print_error = True
        else:
            continue

    if print_error:
        print("注意！由于有些数据的缺失，部分显示的数字可能不会准确！")

    return print_result

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
            try:
                drink_url = fernet_decrypt(str(k_dict["box_drink_in_out_url"]), fernet_key).strip()
                drink_url += "htmlview"

                drink_df = pd.read_html(drink_url, encoding="utf-8")[1]
                parseGoogleHTMLSheet(drink_df)
                print("已通过仅读模式在网络获取酒水出入库数据。")
                print("注意在没有机器人的帮助下，所有的酒水名字和出入库数据不能被自动重置。")

            except Exception as e:
                print()
                print("网络仅读模式获取酒水出入库数据失败, 错误描述如下: ")
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

def parse_sending(payslip_on_duty, drink_on_duty, box_on_duty, cashier_on_duty, google_auth, outlet, send_dict, drink_message_string, tabox_message_string, print_result, k_dict, fernet_key, wifi, date_dict, value_dict, db_writables, backup_foldername, database_url):
    database_url = fernet_decrypt(database_url, fernet_key)

    date = date_dict["dfb"].strftime("%Y-%m-%d")

    drink_stock_alert = eval(k_dict["drink_stock_alert"].strip().capitalize())
    box_stock_alert = eval(k_dict["box_stock_alert"].strip().capitalize())
    night_audit_alert = eval(k_dict["night_audit_alert"].strip().capitalize())
    payslip_end_month_alert = eval(k_dict["payslip_end_month_alert"].strip().capitalize())

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

    write_finance_db = db_writables["write_finance_db"]

    outlet = str(outlet).strip().capitalize()

    rcv_telegram, rcv_email = get_rcv(fernet_key, k_dict, google_auth, backup_foldername, local_database_filename, False)

    if drink_stock_alert:

        drink_stock_email_server = fernet_decrypt(k_dict["drink_stock_email_server"], fernet_key)
        drink_stock_email_sender = fernet_decrypt(k_dict["drink_stock_email_sender"], fernet_key)
        drink_stock_sender_password = fernet_decrypt(k_dict["drink_stock_sender_password"], fernet_key)
        drink_stock_telegram_bot_api = fernet_decrypt(k_dict["drink_stock_telegram_bot_api"], fernet_key)

        if send_drink_msg:
            if len(drink_message_string) > 0:
                if drink_send_channel.strip().capitalize() == "Telegram":
                    receivers = rcv_telegram[drink_on_duty]
                    sending_telegram(is_pr=False,
                                      message=drink_message_string,
                                      api = drink_stock_telegram_bot_api,
                                      receiver=receivers,
                                      wifi=wifi)
                elif drink_send_channel.strip().capitalize() == "Email":
                    receivers = rcv_email[drink_on_duty]
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
                    receivers = rcv_telegram[box_on_duty]
                    sending_telegram(is_pr=False,
                                      message=tabox_message_string,
                                      api = box_stock_telegram_bot_api,
                                      receiver=receivers,
                                      wifi=wifi)
                elif tabox_send_channel.strip().capitalize() == "Email":
                    receivers = rcv_email[box_on_duty]
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
                    receivers = rcv_telegram[cashier_on_duty]
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

                elif night_audit_send_channel.strip().capitalize() == "Email":
                    print_result += ["——————————————", "服务费: ${}".format(svc),
                                     "GST: ${}".format(gst), "日均营业额: ${}".format(ads)]
                    receivers = rcv_email[cashier_on_duty]
                    sending_email(is_pr=True,
                                  mail_server=night_audit_email_server,
                                  mail_sender=night_audit_email_sender,
                                  mail_sender_password=night_audit_sender_password,
                                  mail_receivers=receivers,
                                  mail_subject="{}的报表信息{}".format(outlet, date),
                                  message_string = print_result,
                                  wifi = wifi)
                else:
                    print("Night audit alert sending channel is not defined correctly")
        else:
            if wifi:
                if write_finance_db:
                    if len(print_result) > 0:
                        print("是否重新发送关帐报表? ")
                        action_req = option_num(["重新发送", "不发送"])
                        time.sleep(0.25)
                        user_input = option_limit(action_req, input("在这里输入>>>: "))

                        if user_input == 0:
                            if night_audit_send_channel.strip().capitalize() == "Telegram":
                                receivers = rcv_telegram[cashier_on_duty]
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

                            elif night_audit_send_channel.strip().capitalize() == "Email":
                                print_result += ["——————————————", "服务费: ${}".format(svc),
                                                 "GST: ${}".format(gst), "日均营业额: ${}".format(ads)]

                                receivers = rcv_email[cashier_on_duty]

                                sending_email(is_pr=True,
                                              mail_server=night_audit_email_server,
                                              mail_sender=night_audit_email_sender,
                                              mail_sender_password=night_audit_sender_password,
                                              mail_receivers=receivers,
                                              mail_subject="{}的报表信息{}".format(outlet, date),
                                              message_string = print_result,
                                              wifi = wifi)
                            else:
                                print("Night audit alert sending channel is not defined correctly")
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
                df = pd.read_html(shiftURL+"htmlview", encoding="utf-8")[2]
                parseGoogleHTMLSheet(df)
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
                try:
                    tabox_url = box_drink_in_out_url + "htmlview"
                    tabox_df = pd.read_html(tabox_url, encoding="utf-8")[0]
                    parseGoogleHTMLSheet(tabox_df)
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
                try:
                    drink_url = box_drink_in_out_url + "htmlview"
                    drink_df = pd.read_html(drink_url, encoding="utf-8")[1]
                    parseGoogleHTMLSheet(drink_df)
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

def night_audit_main(database_url, db_setting_url, serialized_rule_filename, service_filename, constants_sheetname, google_auth, box_num, drink_num, promo_num, lun_sales, lun_gc, tb_sales, tb_gc, lun_fwc, lun_kwc, tb_fwc, tb_kwc, night_fwc, night_kwc, script_backup_filename, script, wifi, backup_foldername,cashier_on_duty, drink_on_duty, box_on_duty, payslip_on_duty):
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
                print_result = parse_print_rule(value_dict, rule_df_dict, write_finance_db, write_promo_db, write_drink_db, tabox_write_db)
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
                parse_sending(payslip_on_duty, drink_on_duty, box_on_duty, cashier_on_duty, google_auth, outlet, send_dict, drink_message_string, tabox_message_string, print_result, k_dict, fernet_key, wifi, date_dict, value_dict, db_writables, backup_foldername, database_url)
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

    formURL += "htmlview"

    if df_name == "breakages_df":
        df = pd.read_html(formURL, encoding="utf-8")[index_list_dict[df_name]]
        df.drop("Unnamed: 0", axis=1, inplace=True)
        df.columns = df.iloc[0,:]
        df.drop(0, axis=0, inplace=True)

        df = df[df["选择操作"] == "录入破损物品"]
        df.drop(drop_list_dict[df_name], axis=1, inplace=True)
        df.columns = column_list_dict[df_name]
        df["日期(年年年年-月月-日日)"] = df["日期(年年年年-月月-日日)"].astype(str)
        df["日期(年年年年-月月-日日)"] = df["日期(年年年年-月月-日日)"].apply(lambda x: dt.datetime.strptime(x, "%m/%d/%Y").strftime("%Y-%m-%d"))
        df["日期(年年年年-月月-日日)"] = pd.to_datetime(df["日期(年年年年-月月-日日)"])

        df.reset_index(inplace=True)
        df.drop("index", axis=1, inplace=True)

    elif df_name == "buyInStockDf":
        df = pd.read_html(formURL, encoding="utf-8")[index_list_dict[df_name]]
        df.drop("Unnamed: 0", axis=1, inplace=True)
        df.columns = df.iloc[0,:]
        df.drop(0, axis=0, inplace=True)
        df.reset_index(inplace=True)
        df.drop("index", axis=1, inplace=True)

        df = df[df["选择操作"]== "物品进货"]
        df.drop(drop_list_dict[df_name], axis=1, inplace=True)
        df.columns = column_list_dict[df_name]
        df["进货日期(年年年年-月月-日日)"] = df["进货日期(年年年年-月月-日日)"].astype(str)
        df["进货日期(年年年年-月月-日日)"] = df["进货日期(年年年年-月月-日日)"].apply(lambda x: dt.datetime.strptime(x, "%m/%d/%Y").strftime("%Y-%m-%d"))
        df["进货日期(年年年年-月月-日日)"] = pd.to_datetime(df["进货日期(年年年年-月月-日日)"])

        df.reset_index(inplace=True)
        df.drop("index", axis=1, inplace=True)

    elif df_name == "lockInStockDf":
        df = pd.read_html(formURL, encoding="utf-8")[index_list_dict[df_name]]
        df.drop("Unnamed: 0", axis=1, inplace=True)
        df.columns = df.iloc[0,:]
        df.drop(0, axis=0, inplace=True)
        df.reset_index(inplace=True)
        df.drop("index", axis=1, inplace=True)

    elif df_name == "stockCountDf":
        df = pd.read_html(formURL, encoding="utf-8")[index_list_dict[df_name]]
        df.drop("Unnamed: 0", axis=1, inplace=True)
        df.columns = df.iloc[0,:]
        df.drop(0, axis=0, inplace=True)
        df.reset_index(inplace=True)
        df.drop("index", axis=1, inplace=True)

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
        df["盘点日期"] = pd.to_datetime(df["盘点日期"])
        df.sort_values(by=["盘点日期"], ascending=True, ignore_index=True, inplace=True)

    elif df_name in ["pageOneUnitPrice", "pageTwoUnitPrice", "pageThreeUnitPrice"]:
        df = pd.read_html(formURL, encoding="utf-8")[index_list_dict[df_name]]
        df.drop("Unnamed: 0", axis=1, inplace=True)
        df.columns = df.iloc[0,:]
        df.drop(0, axis=0, inplace=True)
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
                                            try:
                                                box_drink_html_url = box_drink_in_out_url + "htmlview"
                                                tabox_stock_df = pd.read_html(box_drink_html_url, encoding="utf-8")[0]
                                                parseGoogleHTMLSheet(tabox_stock_df)

                                                drink_stock_df = pd.read_html(box_drink_html_url, encoding="utf-8")[1]
                                                parseGoogleHTMLSheet(drink_stock_df)
                                                print("已通过网络仅读模式获取打包盒和酒水出入库信息")
                                            except Exception as e:
                                                print()
                                                print("通过网络仅读模式获取打包盒和酒水出入库信息也失败了")
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
                        df = pd.read_html(shiftURL+"htmlview", encoding="utf-8")
                        pbar1.update(16)

                        pbar1.set_description("读取排班记录")
                        shift_df = df[0]
                        parseGoogleHTMLSheet(shift_df)
                        shift_df["DATE"] = pd.to_datetime(shift_df["DATE"])
                        shift_df["SHIFT"] = shift_df["SHIFT"].astype(str)
                        shift_df["ID"] = shift_df["ID"].astype(int)
                        shift_df["ID"] = shift_df["ID"].astype(str)
                        pbar1.update(16)

                        pbar1.set_description("读取员工资料")
                        employee_info = df[2]
                        parseGoogleHTMLSheet(employee_info)
                        employee_info["ID"] = employee_info["ID"].astype(int)
                        employee_info["ID"] = employee_info["ID"].astype(str)
                        employee_info["FIRST DAY DATE"] = pd.to_datetime(employee_info["FIRST DAY DATE"])
                        employee_info["AL START DATE"] = pd.to_datetime(employee_info["AL START DATE"])
                        employee_info["AL END DATE"] = pd.to_datetime(employee_info["AL END DATE"])
                        pbar1.update(16)

                        pbar1.set_description("读取公共假期")
                        ph_dates_df = df[3]
                        parseGoogleHTMLSheet(ph_dates_df)
                        ph_dates_df["PH DATE"] = pd.to_datetime(ph_dates_df["PH DATE"])
                        pbar1.update(16)

                        pbar1.set_description("读取手动录入假期")
                        leaves_manual_df = df[4]
                        parseGoogleHTMLSheet(leaves_manual_df)
                        leaves_manual_df["DATE"] = pd.to_datetime(leaves_manual_df["DATE"])
                        leaves_manual_df["ID"] = leaves_manual_df["ID"].astype(int)
                        leaves_manual_df["ID"] = leaves_manual_df["ID"].astype(str)

                        pbar1.set_description("读取排班预览")
                        previewSchedule = df[1]
                        previewSchedule.drop("Unnamed: 0", axis=1, inplace=True)
                        previewSchedule = previewSchedule.iloc[:20, :]

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
                        isMonday = eval(str(previewSchedule.iloc[1,3]).capitalize())

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
    
        monday = pd.to_datetime(str(previewSchedule.iloc[1,1]))

        employee_id_show = previewSchedule.iloc[4:18, 0].astype(str).tolist()
        
        if "nan" in employee_id_show:
            employee_id_show.remove("nan")
        
        if len(employee_id_show) < 14:
            for _ in range(14-len(employee_id_show)):
                employee_id_show += [" "]
                
        weekRange = [monday]
        for num in np.arange(1,7):
            weekRange += [ pd.to_datetime(monday + dt.timedelta(days=int(num)))]

        pbar.update(15)
        
        workRange = previewSchedule.copy()
        workRange = workRange.iloc[4:18, 0:10]
    
        work_schedule_df = {"序号" : np.arange(1, 15),
                            "ID" : workRange.iloc[:, 0],
                            "姓名" : workRange.iloc[:,1],
                            "Mon" : workRange.iloc[:,2],
                            "Tue" : workRange.iloc[:,3],
                            "Wed" : workRange.iloc[:,4],
                            "Thu" : workRange.iloc[:,5],
                            "Fri" : workRange.iloc[:,6],
                            "Sat" : workRange.iloc[:,7],
                            "Sun" : workRange.iloc[:,8],}
    
        work_schedule_df["ID"] = work_schedule_df["ID"].astype(str)
        
        rest_days = {}
        
        keys = ["OIL", "PH", "AL", "CCL"]
        sunday = weekRange[-1].strftime("%Y-%m-%d")
        for key in keys:
            key_list = []
            for index in range(len(workRange.iloc[:, 0])):
                id = str(workRange.iloc[:, 0].values.astype(str)[index])
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
        work_schedule_df["SIGN"] = np.repeat(" ", len(workRange))
        work_schedule_df["RE"] = workRange.iloc[:, 9]
        work_schedule_df.drop("ID", axis=1, inplace=True)
        
        pbar.update(15)

        workerCountRange = previewSchedule.copy()
        workerCountRange = workerCountRange.iloc[18:, :9]
        
        for num in range(6):
            workerCountRange["{}".format(num+1)] = np.repeat(" ", len(workerCountRange))

        pbar.update(15)

        weekRemark = str(previewSchedule.iloc[18,9])

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
        
                        table1.add(borb_Paragraph(text, horizontal_alignment=borb_align.CENTERED, font_color=borb_HexColor("#00308F")))
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

def rtn_datetime_column_convert(df):
    df_copy = df.copy()
    time_list = ["日期", "时间"]
    for item in time_list:
        for column in df.columns:
            if item in column:
                df_copy[column] = pd.to_datetime(df_copy[column])
            else:
                pass
    
    return df_copy

def rtn_datetime_strftime(df):
    df_copy = df.copy()
    time_list = ["日期", "时间"]
    for item in time_list:
        for column in df.columns:
            if item in column:
                df_copy[column] = pd.to_datetime(df_copy[column])
                df_copy[column] = df_copy[column].apply(lambda x : x.strftime("%Y-%m-%d %H:%M"))
            else:
                pass
    
    return df_copy

def rtn_get_controls(google_auth, fernet_key, outlet, rtn_control_url, constants_sheetname):
    rtn_control_url = fernet_decrypt(rtn_control_url, fernet_key)

    try:
        rtn_control_sheet = google_auth.open_by_url(rtn_control_url)

        constant_sheet_index = rtn_control_sheet.worksheet(property="title", value=constants_sheetname).index
        rtn_constants = rtn_control_sheet[constant_sheet_index].get_as_df()

        outlet_col_index = rtn_constants.columns.get_loc(key=str(outlet).strip().upper())

        #rtn_constants
        rtn_constants_dict = {}
        for index in range(len(rtn_constants)):
            rtn_constants_dict.update({ str(rtn_constants.iloc[index, 0]) : str(rtn_constants.iloc[index, outlet_col_index]) })
        
        #other controls
        other_controls = {}
        title_list = ["tc", "acm", "sm"]
        sheetname_list = [
                          rtn_constants_dict["table_control_sheetname"],
                          rtn_constants_dict["ala_carte_menu_sheetname"],
                          rtn_constants_dict["set_menu_sheetname"]
                         ]
        
        for item in range(len(title_list)):
            the_sheetname = sheetname_list[item]
            the_index = rtn_control_sheet.worksheet(property="title", value= the_sheetname).index
            other_controls.update({ title_list[item] : rtn_control_sheet[the_index].get_as_df() })

    except Exception as e:
        print()
        print("获取预订初始化信息失败, 错误描述如下: ")
        print()
        print(e)
        rtn_constants_dict = {}
        other_controls = {}
    
    return rtn_constants_dict, other_controls

def rtn_confirm_input(userInput):
    options = option_num(["确定'{}'".format(userInput), "重新输入/不确定/退出"])
    time.sleep(0.25)
    select = option_limit(options, input("在这里输入>>>: "))

    if select == 0:
        return True
    else:
        return False

def rtn_confirm_save(userInput):
    options = option_num(["确定保存'{}'".format(userInput), "重新来过", "不保存/退出"])
    time.sleep(0.25)
    select = option_limit(options, input("在这里输入>>>: "))

    if select == 0:
        confirm_save = True
        end_record_loop = True
    elif select == 1:
        confirm_save = False
        end_record_loop = False
    else:
        confirm_save = False
        end_record_loop = True
    
    return confirm_save, end_record_loop

def rtn_edit_food_item(foodName, rtn_constants_dict, food_items=""):
    print()
    print("正在编辑{}内的菜肴".format(foodName))
    print()
    if food_items == "":
        foodItem = []
    else:
        foodItem = food_items.split(",")
    
    userInput = 0
    while userInput != 3:
        if len(foodItem) > 0:
            for item in foodItem:
                print(item)
        
        options = option_num(["添加菜肴", "编辑菜肴", "删除菜肴", "我已完成编辑菜肴"])
        time.sleep(0.25)
        userInput = option_limit(options, input("在这里输入>>>: "))

        if userInput == 0:
            confirm_foodItemName = False
            while not confirm_foodItemName:
                foodItemName = rtn_input_validation(rule="foodItemName", title="菜肴名", space_removal=True, rtn_constants_dict=rtn_constants_dict)
                confirm_foodItemName = rtn_confirm_input(foodItemName)
            
            foodItem += [foodItemName]
        
        elif userInput == 1:
            if len(foodItem) <= 0:
                print("没有可编辑的菜肴。")
            
            else:
                edit_foodItem = []
                for index in range(len(foodItem)):
                    edit_foodItem += ["编辑'{}'".format(foodItem[index])]

                edit_options = option_num(edit_foodItem)
                time.sleep(0.25)
                editUserInput = option_limit(edit_options, input("在这里输入>>>: "))

                confirm_edit_foodItemName = False
                while not confirm_edit_foodItemName:
                    print("正在编辑{}内的菜肴'{}'".format(foodName, foodItem[editUserInput]))
                    edit_foodItemName = rtn_input_validation(rule="foodItemName", title="菜肴名", space_removal=True, rtn_constants_dict=rtn_constants_dict)
                    confirm_edit_foodItemName = rtn_confirm_input(edit_foodItemName)
                
                foodItem[editUserInput] = edit_foodItemName
        
        elif userInput == 2:
            if len(foodItem) <= 0:
                print("没有可删除的菜肴。")
            
            else:
                remove_foodItem = []
                for index in range(len(foodItem)):
                    remove_foodItem += ["删除'{}'".format(foodItem[index])]
                
                remove_options = option_num(remove_foodItem)
                time.sleep(0.25)
                removeUserInput = option_limit(remove_options, input("在这里输入>>>: "))

                print("确定删除{}内的菜肴'{}'吗? '".format(foodName, foodItem[removeUserInput]))
                confirm_remove_foodItem = rtn_confirm_input(foodItem[removeUserInput])

                if confirm_remove_foodItem:
                    removeItem = foodItem[removeUserInput]
                    foodItem.remove(removeItem)
                    print("已删除'{}'内的菜肴'{}'".format(foodName, removeItem))
                else:
                    print("好的，没有删除。")

    else:
        returnItem = ""

        if len(foodItem) > 0:
            for index in range(len(foodItem)):
                if index != len(foodItem)-1:
                    returnItem += "{},".format(foodItem[index])
                else:
                    returnItem += "{}".format(foodItem[index])

        return returnItem

def rtn_input_validation(rule, title, space_removal=False, ala_carte=False, rtn_constants_dict=None, other_controls=None, table_minimum_pax=None):
    if rule == "searchString":
        maximum_len = int(rtn_constants_dict["max_string_len_allowable_for_search"])
        minimum_len = int(rtn_constants_dict["min_string_len_allowable_for_search"])

        len_test = False
        while not len_test:
            userInput = input("请输入{}: ".format(title))

            if space_removal:
                userInput = remove_spaces(userInput)
                print("空格已自动移除。")
            else:
                pass
            
            if len(userInput) < minimum_len:
                print("输入了{}串字符, 最少输入{}串字符, 请重新输入! ".format(len(userInput), minimum_len))
                len_test = False
            
            else:
                if len(userInput) > maximum_len:
                    print("输入了{}串字符, 最多输入{}串字符, 请重新输入! ".format(len(userInput), maximum_len))
                    len_test = False
                else:
                    len_test = True
        else:
            return str(userInput)

    elif rule == "float":
        input_validate = False
        while not input_validate:
            userInput = input("请输入{}: ".format(title))

            if float_check(userInput):
                input_validate = True
            
            else:
                print("输入的'{}'无效, 必须是数字, 请重新输入! ".format(userInput))
                input_validate = False
        else:
            return float(userInput)

    elif rule == "fqc":
        fqc_max_limit = int(rtn_constants_dict["food_quantity_control_max_limit"])

        input_validate = False
        while not input_validate:
            userInput = input("请输入{}: ".format(title))

            if integer_check(userInput):
                if int(userInput) <= 0:
                    print("输入的'{}'无效, 必须是正整数, 请重新输入! ".format(userInput))
                    input_validate = False
                else:
                    if int(userInput) <= fqc_max_limit:
                        input_validate = True
                    else:
                        print("输入的'{}'无效, 请重新输入! ".format(userInput))
                        input_validate = False
            else:
                print("输入的字符无效, 必须是正整数, 请重新输入! ".format(userInput))
                input_validate = False
        else:
            return int(userInput)

    elif rule == "foodName":
        if ala_carte:
            foodName_allow_dup = eval(str(rtn_constants_dict["acm_foodName_allow_duplicate"]).strip().capitalize())
            menu = other_controls["acm"].copy()
            menu["菜名"] = menu["菜名"].astype(str)
        else:
            foodName_allow_dup = eval(str(rtn_constants_dict["sm_foodName_allow_duplicate"]).strip().capitalize())
            menu = other_controls["sm"].copy()
            menu["套餐名"] = menu["套餐名"].astype(str)

        max_len = int(rtn_constants_dict["max_string_len_allowable_for_foodName"])
        min_len = int(rtn_constants_dict["min_string_len_allowable_for_foodName"])

        foodName_dup_test = False
        len_test = False
        while not len_test or not foodName_dup_test:
            userInput = input("请输入{}: ".format(title))

            if space_removal:
                userInput = remove_spaces(userInput)
            else:
                pass

            if len(userInput) > max_len:
                print("输入了{}串字符, 最多输入{}串字符, 请重新输入! ".format(len(userInput), len(max_len)))
                len_test = False
            else:
                if len(userInput) < min_len:
                    print("输入了{}串字符, 最少输入{}串字符, 请重新输入！".format(len(userInput), len(min_len)))
                    len_test = False
                else:
                    len_test = True
            
            if not foodName_allow_dup:
                if ala_carte:
                    emptiness = menu[menu["菜名"] == userInput].empty
                else:
                    emptiness = menu[menu["套餐名"] == userInput].empty
                
                if not emptiness:
                    print("菜名/套餐名'{}'重复, 请重新输入! ".format(userInput))
                    foodName_dup_test = False
                else:
                    foodName_dup_test = True
            else:
                foodName_dup_test = True
        else:
            return str(userInput)

    elif rule == "foodPrice":
        min_price = float(rtn_constants_dict["food_min_price"])
        max_price = float(rtn_constants_dict["food_max_price"])

        price_test = False
        while not price_test:
            userInput = input("请输入{}".format(title))

            if float_check(userInput):
                if float(userInput) < min_price:
                    print("输入的'{}'无效, 不能小于'{}', 请重新输入! ".format(userInput, min_price))
                    price_test = False
                else:
                    if float(userInput) > max_price:
                        print("输入的'{}'无效, 不能大于'{}', 请重新输入! ".format(userInput, max_price))
                        price_test = False
                    else:
                        price_test = True
            else:
                print("输入的字符无效, 必须是数字不小于{}不大于{}, 请重新输入! ".format(min_price, max_price))
                price_test = False
        else:
            return float(format(float(userInput), ".2f"))

    elif rule == "foodItemName":
        maximum_len = int(rtn_constants_dict["max_string_len_allowable_for_foodName"])
        minimum_len = int(rtn_constants_dict["min_string_len_allowable_for_foodName"])

        len_test = False
        while not len_test:
            userInput = input("请输入{}: ".format(title))

            if space_removal:
                userInput = remove_spaces(userInput)
                print("空格已自动移除。")
            else:
                pass
            
            if len(userInput) < minimum_len:
                print("输入了{}串字符, 最少输入{}串字符, 请重新输入! ".format(len(userInput), minimum_len))
                len_test = False
            
            else:
                if len(userInput) > maximum_len:
                    print("输入了{}串字符, 最多输入{}串字符, 请重新输入! ".format(len(userInput), maximum_len))
                    len_test = False
                else:
                    len_test = True
        else:
            return str(userInput)

    elif rule == "tableName":
        tc = other_controls["tc"]
        tc["桌名"] = tc["桌名"].astype(str)

        max_len = int(rtn_constants_dict["tableName_max_len"])
        min_len = int(rtn_constants_dict["tableName_min_len"])

        len_test = False
        while not len_test:
            userInput = input("请输入{}: ".format(title))

            if space_removal:
                userInput = remove_spaces(userInput)
            else:
                pass

            if len(userInput) > max_len:
                print("输入了{}串字符, 最多输入{}串字符, 请重新输入! ".format(len(userInput), max_len))
                len_test = False
            else:
                if len(userInput) < min_len:
                    print("输入了{}串字符, 最少输入{}串字符, 请重新输入！".format(len(userInput), min_len))
                    len_test = False
                else:
                    len_test = True
        else:
            return str(userInput)

    elif rule == "tableMin":
        table_min = int(rtn_constants_dict["tc_minimum_pax"])
        table_max = int(rtn_constants_dict["tc_maximum_pax"])

        input_validate = False
        while not input_validate:
            userInput = input("请输入{}: ".format(title))

            if integer_check(userInput):
                if int(userInput) < table_min:
                    print("输入的'{}'无效, 不能小于{}, 请重新输入! ".format(userInput, table_min))
                    input_validate = False
                else:
                    if int(userInput) > table_max:
                        print("输入的'{}'无效, 不能大于'{}', 请重新输入! ".format(userInput, table_max))
                        input_validate = False
                    else:
                        input_validate = True
            else:
                print("输入的字符无效，必须是整数不小于{}不大于{}, 请重新输入! ".format(table_min, table_max))
                input_validate = False
        else:
            return int(userInput)

    elif rule == "tableMax":
        table_min = table_minimum_pax
        table_max = int(rtn_constants_dict["tc_maximum_pax"])

        input_validate = False
        while not input_validate:
            userInput = input("请输入{}: ".format(title))

            if integer_check(userInput):
                if int(userInput) < table_min:
                    print("输入的'{}'无效, 不能小于{}, 请重新输入! ".format(userInput, table_min))
                    input_validate = False
                else:
                    if int(userInput) > table_max:
                        print("输入的'{}'无效, 不能大于'{}', 请重新输入! ".format(userInput, table_max))
                        input_validate = False
                    else:
                        input_validate = True
            else:
                print("输入的字符无效，必须是整数不小于{}不大于{}, 请重新输入! ".format(table_min, table_max))
                input_validate = False
        else:
            return int(userInput)

    elif rule == "orderNumber":
        maximum_len = int(rtn_constants_dict["orderNumber_max_len"])
        minimum_len = int(rtn_constants_dict["orderNumber_min_len"])

        len_test = False
        while not len_test:
            userInput = input("请输入{}: ".format(title))

            if space_removal:
                userInput = remove_spaces(userInput)
                print("空格已自动移除。")
            else:
                pass
            
            if len(userInput) < minimum_len:
                print("输入了{}串字符, 最少输入{}串字符, 请重新输入! ".format(len(userInput), minimum_len))
                len_test = False
            
            else:
                if len(userInput) > maximum_len:
                    print("输入了{}串字符, 最多输入{}串字符, 请重新输入! ".format(len(userInput), maximum_len))
                    len_test = False
                else:
                    len_test = True
        else:
            return str(userInput)

    elif rule == "localPhone":
        max_value = 99999999
        min_value = 30000000

        input_validate = False
        while not input_validate:
            userInput = input("请输入{}: ".format(title))

            if integer_check(userInput):
                if int(userInput) < min_value:
                    print("输入的'{}'无效, 必须是8位数以3,6,8,9开头".format(userInput))
                    print("请重新输入! ")
                    input_validate = False
                else:
                    if int(userInput) > max_value:
                        print("输入的'{}'无效, 必须是8位数以3,6,8,9开头".format(userInput))
                        print("请重新输入! ")
                        input_validate = False
                    else:
                        if str(userInput)[0] not in ["3", "6", "8", "9"]:
                            print("输入的'{}'无效, 必须是8位数以3,6,8,9开头".format(userInput))
                            print("请重新输入! ")
                            input_validate = False
                        else:
                            input_validate = True
            else:
                print("输入的字符无效, 必须是8位数以3,6,8,9开头。")
                print("请重新输入! ")
                input_validate = False
        else:
            return str(int(userInput))

    elif rule == "pax":
        tc = other_controls["tc"]
        tc["最小载客量"] = tc["最小载客量"].astype(int)
        tc["最大载客量"] = tc["最大载客量"].astype(int)

        try:
            max_value = max(tc["最大载客量"])
            min_value = min(tc["最小载客量"])
        except:
            max_value = 1
            min_value = 1

        input_validate = False
        while not input_validate:
            userInput = input("请输入{}: ".format(title))

            if integer_check(userInput):
                if int(userInput) < min_value:
                    print("输入的'{}'无效, 必须是正整数不小于{}不大于{}。".format(userInput, min_value, max_value))
                    print("请重新输入! ")
                    input_validate = False
                else:
                    if int(userInput) > max_value:
                        print("输入的'{}'无效, 必须是正整数不小于{}不大于{}。".format(userInput, min_value, max_value))
                        print("请重新输入! ")
                        input_validate = False
                    else:
                        input_validate = True
            else:
                print("输入的字符无效, 必须是正整数不小于{}不大于{}。".format(min_value, max_value))
                print("请重新输入! ")
        else:
            return int(userInput)

    elif rule == "remark":
        maximum_len = int(rtn_constants_dict["max_string_len_allowable_for_remark"])
        minimum_len = 0

        len_test = False
        while not len_test:
            userInput = input("请输入{}: ".format(title))

            if space_removal:
                userInput = remove_spaces(userInput)
                print("空格已自动移除。")
            else:
                pass
            
            if len(userInput) < minimum_len:
                print("输入了{}串字符, 最少输入{}串字符, 请重新输入! ".format(len(userInput), minimum_len))
                len_test = False
            
            else:
                if len(userInput) > maximum_len:
                    print("输入了{}串字符, 最多输入{}串字符, 请重新输入! ".format(len(userInput), maximum_len))
                    len_test = False
                else:
                    len_test = True
        else:
            return str(userInput)

    elif rule == "foodQty":
        table_min = int(rtn_constants_dict["food_quantity_min"])
        table_max = int(rtn_constants_dict["food_quantity_max"])

        input_validate = False
        while not input_validate:
            userInput = input("请输入{}: ".format(title))

            if integer_check(userInput):
                if int(userInput) < table_min:
                    print("输入的'{}'无效, 不能小于{}, 请重新输入! ".format(userInput, table_min))
                    input_validate = False
                else:
                    if int(userInput) > table_max:
                        print("输入的'{}'无效, 不能大于'{}', 请重新输入! ".format(userInput, table_max))
                        input_validate = False
                    else:
                        input_validate = True
            else:
                print("输入的字符无效，必须是整数不小于{}不大于{}, 请重新输入! ".format(table_min, table_max))
                input_validate = False
        else:
            return int(userInput)
    
    elif rule == "notAdultPax":
        tc = other_controls["tc"]
        tc["最小载客量"] = tc["最小载客量"].astype(int)
        tc["最大载客量"] = tc["最大载客量"].astype(int)

        try:
            max_value = max(tc["最大载客量"])
            min_value = 0
        except:
            max_value = 1
            min_value = 0

        input_validate = False
        while not input_validate:
            userInput = input("请输入{}: ".format(title))

            if integer_check(userInput):
                if int(userInput) < min_value:
                    print("输入的'{}'无效, 必须是正整数不小于{}不大于{}。".format(userInput, min_value, max_value))
                    print("请重新输入! ")
                    input_validate = False
                else:
                    if int(userInput) > max_value:
                        print("输入的'{}'无效, 必须是正整数不小于{}不大于{}。".format(userInput, min_value, max_value))
                        print("请重新输入! ")
                        input_validate = False
                    else:
                        input_validate = True
            else:
                print("输入的字符无效, 必须是正整数不小于{}不大于{}。".format(min_value, max_value))
                print("请重新输入! ")
        else:
            return int(userInput)

def rtn_select_table(rtn_constants_dict, other_controls):
    tc = other_controls["tc"]
    tc["桌名"] = tc["桌名"].astype(str)
    tc["最小载客量"] = tc["最小载客量"].astype(int)
    tc["最大载客量"] = tc["最大载客量"].astype(int)

    slice_tc = tc.copy()
    slice_tc["桌名"] = slice_tc["桌名"].astype(str)
    slice_tc["最小载客量"] = slice_tc["最小载客量"].astype(int)
    slice_tc["最大载客量"] = slice_tc["最大载客量"].astype(int)

    if tc.empty:
        print("无桌台! ")
        choose_table_input = None

    else:
        options = option_num(["按桌名搜索", "按桌子ID搜索", "按照规则重新排列所有桌台搜索", "退出桌台选取"])
        time.sleep(0.25)
        userInput = option_limit(options, input("在这里输入>>>: "))

        if userInput in [0, 1]:
            if userInput == 0:
                search_string_title = "桌名"
            else:
                search_string_title = "桌子ID"
            
            search_string = rtn_input_validation(rule="searchString", title=search_string_title, space_removal=True, rtn_constants_dict=rtn_constants_dict)

            if userInput == 0:
                slice_tc = slice_tc[slice_tc["桌名"].str.contains(search_string)]
            
            else:
                slice_tc = slice_tc[slice_tc["桌子ID"].str.contains(search_string)]
        
        elif userInput == 2:
            filterOptions = option_num([
                "按桌名从大到小排列",
                "按桌名从小到大排列",
                "按最小载客量从大到小排列",
                "按最小载客量从小到大排列",
                "按最大载客量从大到小排列",
                "按最大载客量从小到大排列",
                "按桌子ID从大到小排列",
                "按桌子ID从小到大排列",
            ])

            time.sleep(0.25)
            filter_select = option_limit(filterOptions, input("在这里输入>>>: "))

            if filter_select == 0:
                slice_tc.sort_values(by="桌名", ascending=False, inplace=True, ignore_index=True)
            
            elif filter_select == 1:
                slice_tc.sort_values(by="桌名", ascending=True, inplace=True, ignore_index=True)
            
            elif filter_select == 2:
                slice_tc.sort_values(by="最小载客量", ascending=False, inplace=True, ignore_index=True)
            
            elif filter_select == 3:
                slice_tc.sort_values(by="最小载客量", ascending=True, inplace=True, ignore_index=True)
            
            elif filter_select == 4:
                slice_tc.sort_values(by="最大载客量", ascending=False, inplace=True, ignore_index=True)

            elif filter_select == 5:
                slice_tc.sort_values(by="最大载客量", ascending=True, inplace=True, ignore_index=True)
            
            elif filter_select == 6:
                slice_tc.sort_values(by="桌子ID", ascending=False, inplace=True, ignore_index=True)
            
            elif filter_select == 7:
                slice_tc.sort_values(by="桌子ID", ascending=True, inplace=True, ignore_index=True)
        
        elif userInput == 3:
            print("好的, 已退出桌台选取。")
            choose_table_input = None
        
        if userInput != 3:
            if slice_tc.empty:
                print("未找到任何匹配的桌台。")
                choose_table_input = None
            else:
                slice_tc.reset_index(inplace=True)
                slice_tc.drop("index", axis=1, inplace=True)

                slice_tc_list = []

                for index in range(len(slice_tc)):
                    slice_tc_list += ["'桌名: {}, {}-{}人桌'".format(slice_tc.iloc[index, 1], slice_tc.iloc[index, 2], slice_tc.iloc[index, 3])]
                
                choose_table_options = option_num(slice_tc_list)
                time.sleep(0.25)
                choose_table_input = option_limit(choose_table_options, input("在这里输入>>>: "))
        else:
            choose_table_input = None
    
    if isinstance(choose_table_input, int):
        return str(slice_tc.iloc[choose_table_input, 0])
    else:
        return None

def rtn_select_food(ala_carte, rtn_constants_dict, other_controls):
    #select respective menu and done data type conversion
    if ala_carte:
        menu = other_controls["acm"]
        slice_menu = menu.copy()
        menu["菜名"] = menu["菜名"].astype(str)
        slice_menu["菜名"] = slice_menu["菜名"].astype(str)

    else:
        menu = other_controls["sm"]
        slice_menu = menu.copy()
        menu["套餐名"] = menu["套餐名"].astype(str)
        menu["菜肴"] = menu["菜肴"].astype(str)
        slice_menu["套餐名"] = slice_menu["套餐名"].astype(str)
        slice_menu["菜肴"] = slice_menu["菜肴"].astype(str)

    menu["菜品ID"] = menu["菜品ID"].astype(str)
    menu["价格"] = menu["价格"].astype(float)
    slice_menu["菜品ID"] = slice_menu["菜品ID"].astype(str)
    slice_menu["价格"] = slice_menu["价格"].astype(float)

    if menu.empty:
        #empty menu return None
        print("菜单空白! ")
        choose_food_input = None

    else:
        options = option_num(["按菜名/套餐名搜索", "按价格搜索", "按菜品ID搜索", "退出菜品选取"])
        time.sleep(0.25)
        userInput = option_limit(options, input("在这里输入>>>: "))

        if userInput in [0, 2]:
            if userInput == 0:
                search_string_title = "菜名/套餐名"
            else:
                search_string_title = "菜品ID"

            search_string = rtn_input_validation(rule="searchString", title=search_string_title, space_removal=True, rtn_constants_dict=rtn_constants_dict)
            
            if userInput == 0:
                if ala_carte:
                    slice_menu = slice_menu[slice_menu["菜名"].str.contains(search_string)]
                else:
                    slice_menu = slice_menu[slice_menu["套餐名"].str.contains(search_string)]
            else:
                slice_menu = slice_menu[slice_menu["菜品ID"].str.contains(search_string)]
        
        elif userInput == 1:
            float_value = rtn_input_validation(rule="float", title="价格")
            slice_menu = slice_menu[slice_menu["价格"] == float_value]
        
        elif userInput == 3:
            print("好的，已退出菜品选取")
            choose_food_input = None

        if userInput != 3:
            if slice_menu.empty:
                print("未找到任何匹配的菜名/套餐名/菜品ID/价格")
                choose_food_input = None

            else:
                slice_menu.reset_index(inplace=True)
                slice_menu.drop("index", axis=1, inplace=True)

                slice_menu_list = []

                for index in range(len(slice_menu)):
                    slice_menu_list += ["'菜名/套餐名: {}, 价格: {} '".format(slice_menu.iloc[index, 1], slice_menu.iloc[index, 2])]
                
                choose_food_options = option_num(slice_menu_list)
                time.sleep(0.25)
                choose_food_input = option_limit(choose_food_options, input("在这里输入>>>: "))
        else:
            choose_food_input = None

    if isinstance(choose_food_input, int):
        return str(slice_menu.iloc[choose_food_input, 0])
    else:
        return None

def rtn_edit_menu(rtn_constants_dict, other_controls):
    acm = other_controls["acm"].copy()
    sm = other_controls["sm"].copy()

    acm_columns = str(rtn_constants_dict["acm_columns"]).split(",")
    sm_columns = str(rtn_constants_dict["sm_columns"]).split(",")

    acm["菜品ID"] = acm["菜品ID"].astype(str)
    sm["菜品ID"] = sm["菜品ID"].astype(str)

    acm = rtn_datetime_column_convert(acm)
    sm = rtn_datetime_column_convert(sm)

    actionInput = 0
    while actionInput != 6:
        print()
        print()
        if acm.empty:
            print("暂无单点。")
        else:
            print("单点菜单详情: ")
            prtdf(acm)
        print()
        print()
        if sm.empty:
            print("暂无套餐。")
        else:
            print("套餐菜单详情: ")
            for index in range(len(sm)):
                foodName = str(sm.iloc[index, 1])
                foodPrice = format(float(sm.iloc[index, 2]), ".2f")
                food_items = str(sm.iloc[index, 3]).split(",")

                print("{}. {}".format(index+1, foodName))
                print("价格: {}".format(foodPrice))

                for i in range(len(food_items)):
                    print("第{}道菜: {}".format(i+1, food_items[i]))
                
                print()
        print()
        print()

        options = option_num(["添加单点菜品", "编辑单点", "移除单点菜品", "添加套餐", "编辑套餐", "移除套餐", "我已完成编辑菜单"])
        time.sleep(0.25)
        actionInput = option_limit(options, input("在这里输入>>>: "))

        if actionInput == 0:
            end_record_loop = False
            while not end_record_loop:
                #foodId, foodName, Price, time_created
                foodId = str(uuid.uuid1().hex)

                confirm_foodName = False
                while not confirm_foodName:
                    foodName = rtn_input_validation(rule="foodName", title="菜名/套餐名", space_removal=True, ala_carte=True, rtn_constants_dict=rtn_constants_dict, other_controls=other_controls)
                    confirm_foodName = rtn_confirm_input(foodName)
                
                confirm_foodPrice = False
                while not confirm_foodPrice:
                    foodPrice = rtn_input_validation(rule="foodPrice", title="价格", rtn_constants_dict=rtn_constants_dict)
                    confirm_foodPrice = rtn_confirm_input(foodPrice)
                
                time_created = pd.to_datetime(dt.datetime.now())
                sequencial_list = [foodId, foodName, foodPrice, time_created]

                new_ac_food = {}
                for index in range(len(acm_columns)):
                    new_ac_food.update({ acm_columns[index] : [sequencial_list[index]] })
                
                new_ac_food = pd.DataFrame(new_ac_food)

                confirm_save, end_record_loop = rtn_confirm_save(foodName)
            else:
                if confirm_save:
                    ac_concat = pd.concat([acm, new_ac_food], ignore_index=True)
                    ac_concat.sort_values(by="创建时间", ascending=False, inplace=True, ignore_index=True)
                    acm = ac_concat
                    other_controls["acm"] = ac_concat
                    print("'{}'保存成功! ".format(foodName))
                else:
                    print("'{}'未保存, 已退出。".format(foodName))

        elif actionInput == 1:
            foodId = rtn_select_food(ala_carte=True, rtn_constants_dict=rtn_constants_dict, other_controls=other_controls)

            if isinstance(foodId, str):
                old_foodName = str(acm[acm["菜品ID"] == foodId]["菜名"].values[0])
                old_foodPrice = float(acm[acm["菜品ID"] == foodId]["价格"].values[0])

                edit_options = option_num(["编辑菜名'{}'".format(old_foodName), "编辑其价格({})".format(old_foodPrice), "退出编辑单点"])
                time.sleep(0.25)
                edit_select = option_limit(edit_options, input("在这里输入>>>: "))

                if edit_select == 0:
                    end_record_loop = False
                    while not end_record_loop:
                        foodName = rtn_input_validation(rule="foodName", title="菜名/套餐名", space_removal=True, ala_carte=True, rtn_constants_dict=rtn_constants_dict, other_controls=other_controls)
                        
                        print()
                        print("确定把'{}'换成'{}'吗? ".format(old_foodName, foodName))

                        confirm_save, end_record_loop = rtn_confirm_save(foodName)
                    
                    else:
                        if confirm_save:
                            #foodName = str(acm[acm["菜品ID"] == foodId]["菜名"].values[0])
                            foodPrice = str(acm[acm["菜品ID"] == foodId]["价格"].values[0])
                            foodTimeCreated = pd.to_datetime(str(acm[acm["菜品ID"] == foodId]["创建时间"].values[0]))
                            acm = other_controls["acm"].copy()
                            acm["菜品ID"] = acm["菜品ID"].astype(str)
                            acm = acm[acm["菜品ID"] != foodId]

                            sequence = [
                                foodId,
                                foodName,
                                foodPrice,
                                foodTimeCreated
                            ]

                            modify = {}
                            for index in range(len(acm_columns)):
                                modify.update({ acm_columns[index] : [sequence[index]] })

                            modify = pd.DataFrame(modify)

                            acm = pd.concat([acm, modify], ignore_index=True)
                            acm["创建时间"] = pd.to_datetime(acm["创建时间"])
                            acm.sort_values(by="创建时间", ascending=False, inplace=True, ignore_index=True)

                            other_controls["acm"] = acm

                            print("编辑成功, '{}'已成功换成'{}'! ".format(old_foodName, foodName))
                        else:
                            print("编辑失败, 保留其菜名'{}'。".format(old_foodName))
                    
                elif edit_select == 1:
                    end_record_loop = False
                    while not end_record_loop:
                        foodPrice = rtn_input_validation(rule="foodPrice", title="价格", rtn_constants_dict=rtn_constants_dict)

                        print("菜名: {}".format(foodName))
                        print("确定把其价格从'{}'换成'{}'吗? ".format(format(old_foodPrice, ".2f"), format(foodPrice, ".2f")))

                        confirm_save, end_record_loop = rtn_confirm_save(foodPrice)
                    else:
                        if confirm_save:
                            foodName = str(acm[acm["菜品ID"] == foodId]["菜名"].values[0])
                            #foodPrice = str(acm[acm["菜品ID"] == foodId]["价格"].values[0])
                            foodTimeCreated = pd.to_datetime(str(acm[acm["菜品ID"] == foodId]["创建时间"].values[0]))
                            acm = other_controls["acm"].copy()
                            acm["菜品ID"] = acm["菜品ID"].astype(str)
                            acm = acm[acm["菜品ID"] != foodId]

                            sequence = [
                                foodId,
                                foodName,
                                foodPrice,
                                foodTimeCreated
                            ]

                            modify = {}
                            for index in range(len(acm_columns)):
                                modify.update({ acm_columns[index] : [sequence[index]] })

                            modify = pd.DataFrame(modify)

                            acm = pd.concat([acm, modify], ignore_index=True)
                            acm["创建时间"] = pd.to_datetime(acm["创建时间"])
                            acm.sort_values(by="创建时间", ascending=False, inplace=True, ignore_index=True)

                            other_controls["acm"] = acm

                            print("编辑成功, '{}'已成功换成'{}'! ".format(old_foodPrice, foodPrice))
                        else:
                            print("编辑失败, 保留其价格'{}'。".format(old_foodPrice))

        elif actionInput == 2:
            foodId = rtn_select_food(ala_carte=True, rtn_constants_dict=rtn_constants_dict, other_controls=other_controls)

            if isinstance(foodId, str):
                foodName = str(acm[acm["菜品ID"] == foodId]["菜名"].values[0])

                print("确定删除'{}'吗? ".format(foodName))
                confirm_del = rtn_confirm_input(foodName)

                if confirm_del:
                    acm = other_controls["acm"].copy()
                    acm["菜品ID"] = acm["菜品ID"].astype(str)
                    acm = acm[acm["菜品ID"] != foodId]

                    other_controls["acm"] = acm
                    print("'{}'删除成功。".format(foodName))
                
                else:
                    print("好的，没有删除。")
        
        elif actionInput == 3:
            end_record_loop = False
            while not end_record_loop:
                #foodId, foodName, foodPrice, foodItem time_created
                foodId = str(uuid.uuid1().hex)

                confirm_foodName = False
                while not confirm_foodName:
                    foodName = rtn_input_validation(rule="foodName", title="菜名/套餐名", space_removal=True, ala_carte=False, rtn_constants_dict=rtn_constants_dict, other_controls=other_controls)
                    confirm_foodName = rtn_confirm_input(foodName)

                confirm_foodPrice = False
                while not confirm_foodPrice:
                    foodPrice = rtn_input_validation(rule="foodPrice", title="价格", rtn_constants_dict=rtn_constants_dict)
                    confirm_foodPrice = rtn_confirm_input(foodPrice)

                foodItem = rtn_edit_food_item(foodName=foodName, rtn_constants_dict=rtn_constants_dict, food_items="")
                
                time_created = pd.to_datetime(dt.datetime.now())
                sequencial_list = [foodId, foodName, foodPrice, foodItem, time_created]

                new_sm_food = {}
                for index in range(len(sm_columns)):
                    new_sm_food.update({ sm_columns[index] : [sequencial_list[index]] })
                
                new_sm_food = pd.DataFrame(new_sm_food)
                confirm_save, end_record_loop = rtn_confirm_save(foodName)
            else:
                if confirm_save:
                    sm_concat = pd.concat([sm, new_sm_food], ignore_index=True)
                    sm_concat.sort_values(by="创建时间", ascending=False, inplace=True, ignore_index=True)
                    sm = sm_concat
                    other_controls["sm"] = sm_concat
                    print("'{}'保存成功! ".format(foodName))
                else:
                    print("'{}'未保存，已退出。".format(foodName))

        elif actionInput == 4:
            foodId = rtn_select_food(ala_carte=False, rtn_constants_dict=rtn_constants_dict, other_controls=other_controls)

            if isinstance(foodId, str):
                old_foodName = str(sm[sm["菜品ID"] == foodId]["套餐名"].values[0])
                old_foodPrice = float(sm[sm["菜品ID"] == foodId]["价格"].values[0])
                old_foodItem = str(sm[sm["菜品ID"] == foodId]["菜肴"].values[0])

                edit_options = option_num(["编辑套餐名'{}'".format(old_foodName), "编辑其价格({})".format(old_foodPrice), "编辑其菜肴", "退出编辑套餐"])
                time.sleep(0.25)
                edit_select = option_limit(edit_options, input("在这里输入>>>: "))

                if edit_select == 0:
                    end_record_loop = False
                    while not end_record_loop:
                        foodName = rtn_input_validation(rule="foodName", title="菜名/套餐名", space_removal=False, ala_carte=True, rtn_constants_dict=rtn_constants_dict, other_controls=other_controls)

                        print()
                        print()
                        print("确定把'{}'换成'{}'吗? ".format(old_foodName, foodName))
                        confirm_save, end_record_loop = rtn_confirm_save(foodName)
                    else:
                        if confirm_save:
                            foodPrice = float(format(float(sm[sm["菜品ID"] == foodId]["价格"].values[0]), ".2f"))
                            foodItem = str(sm[sm["菜品ID"] == foodId]["菜肴"].values[0])
                            foodCreatedTime = pd.to_datetime(sm[sm["菜品ID"] == foodId]["创建时间"].values[0])

                            sm = other_controls["sm"].copy()
                            sm["菜品ID"] = sm["菜品ID"].astype(str)
                            sm = sm[sm["菜品ID"] != foodId]

                            sequence = [
                                foodId,
                                foodName,
                                foodPrice,
                                foodItem,
                                foodCreatedTime,
                            ]
                            
                            modify = {}
                            for index in range(len(sm_columns)):
                                modify.update({ sm_columns[index] : [ sequence[index] ] })

                            modify = pd.DataFrame(modify)

                            sm  = pd.concat([sm, modify], ignore_index=True)
                            sm["创建时间"] = pd.to_datetime(sm["创建时间"])
                            sm.sort_values(by="创建时间", ascending=False, ignore_index=True, inplace=True)
                            other_controls["sm"] = sm

                            print("编辑成功, '{}'已成功换成'{}'! ".format(old_foodName, foodName))
                        else:
                            print("编辑失败, 保留其套餐名'{}'。".format(old_foodName))
                
                elif edit_select == 1:
                    end_record_loop = False
                    while not end_record_loop:
                        foodPrice = rtn_input_validation(rule="foodPrice", title="价格", rtn_constants_dict=rtn_constants_dict)

                        print()
                        print()
                        print("套餐名: {}".format(old_foodName))
                        print("确定把其价格从'{}'换成'{}'吗? ".format(format(old_foodPrice, ".2f"), format(foodPrice, ".2f")))

                        confirm_save, end_record_loop = rtn_confirm_save(foodPrice)
                    else:
                        if confirm_save:
                            foodName = str(sm[sm["菜品ID"] == foodId]["套餐名"].values[0])
                            foodItem = str(sm[sm["菜品ID"] == foodId]["菜肴"].values[0])
                            foodCreatedTime = pd.to_datetime(sm[sm["菜品ID"] == foodId]["创建时间"].values[0])

                            sm = other_controls["sm"].copy()
                            sm["菜品ID"] = sm["菜品ID"].astype(str)
                            sm = sm[sm["菜品ID"] != foodId]

                            sequence = [
                                foodId,
                                foodName,
                                foodPrice,
                                foodItem,
                                foodCreatedTime,
                            ]
                            
                            modify = {}
                            for index in range(len(sm_columns)):
                                modify.update({ sm_columns[index] : [ sequence[index] ] })

                            modify = pd.DataFrame(modify)

                            sm  = pd.concat([sm, modify], ignore_index=True)
                            sm["创建时间"] = pd.to_datetime(sm["创建时间"])
                            sm.sort_values(by="创建时间", ascending=False, ignore_index=True, inplace=True)
                            other_controls["sm"] = sm

                            print("编辑成功, 其价格已成功改成'{}'! ".format(format(foodPrice, ".2f")))

                        else:
                            print("编辑失败, 保留其价格'{}'。".format(old_foodPrice))
                            
                elif edit_select == 2:
                    foodItem = rtn_edit_food_item(foodName=foodName, rtn_constants_dict=rtn_constants_dict, food_items=old_foodItem)

                    print("确定保存其修改过的菜肴吗? ")
                    confirm_save_foodItem = rtn_confirm_input("保存其修改过的菜肴")

                    if confirm_save_foodItem:
                        foodName = str(sm[sm["菜品ID"] == foodId]["套餐名"].values[0])
                        foodPrice = float(format(float(sm[sm["菜品ID"] == foodId]["价格"].values[0]), ".2f"))
                        foodCreatedTime = pd.to_datetime(sm[sm["菜品ID"] == foodId]["创建时间"].values[0])

                        sm = other_controls["sm"].copy()
                        sm["菜品ID"] = sm["菜品ID"].astype(str)
                        sm = sm[sm["菜品ID"] != foodId]

                        sequence = [
                            foodId,
                            foodName,
                            foodPrice,
                            foodItem,
                            foodCreatedTime,
                        ]
                        
                        modify = {}
                        for index in range(len(sm_columns)):
                            modify.update({ sm_columns[index] : [ sequence[index] ] })

                        modify = pd.DataFrame(modify)

                        sm  = pd.concat([sm, modify], ignore_index=True)
                        sm["创建时间"] = pd.to_datetime(sm["创建时间"])
                        sm.sort_values(by="创建时间", ascending=False, ignore_index=True, inplace=True)
                        other_controls["sm"] = sm
                        
                        print("套餐'{}'的菜肴已成功修改! ".format(foodName))
                    else:
                        print("套餐'{}'的菜肴未修改, 保留其原来的菜肴! ".format(foodName))

        elif actionInput == 5:
            foodId = rtn_select_food(ala_carte=False, rtn_constants_dict=rtn_constants_dict, other_controls=other_controls)

            if isinstance(foodId, str):
                foodName = str(sm[sm["菜品ID"] == foodId]["套餐名"].values[0])

                print("确定删除'{}'吗? ".format(foodName))
                confirm_del = rtn_confirm_input(foodName)

                if confirm_del:
                    sm = other_controls["sm"].copy()
                    sm["菜品ID"] = sm["菜品ID"].astype(str)

                    sm = sm[sm["菜品ID"] != foodId]
                    other_controls["sm"] = sm

                    print("'{}'删除成功。".format(foodName))
                
                else:
                    print("好的, 没有删除。")

    else:
        return acm, sm, other_controls               

def rtn_edit_tables(rtn_constants_dict, other_controls):
    tc = other_controls["tc"]
    tc["桌子ID"] = tc["桌子ID"].astype(str)

    tc = rtn_datetime_column_convert(tc)

    tc_columns = str(rtn_constants_dict["tc_columns"]).split(",")

    tableName_allow_dup = eval(str(rtn_constants_dict["tableName_allow_dup"]).strip().capitalize())

    tc_select = 0
    while tc_select != 3:
        print()
        print()
        print()
        if tc.empty:
            print("暂无桌台。")
        else:
            print("桌台详情: ")
            prtdf(tc.drop("桌子ID", axis=1))
        print()
        print()
        print()
        tc_options = option_num(["添加桌台", "编辑桌台", "删除桌台", "我已完成编辑桌台"])
        time.sleep(0.25)
        tc_select = option_limit(tc_options, input("在这里输入>>>: "))

        if tc_select == 0:
            end_record_loop = False
            while not end_record_loop:
                confirm_tableName = False
                while not confirm_tableName:
                    tableName = rtn_input_validation(rule="tableName", title="桌名", space_removal=True, rtn_constants_dict=rtn_constants_dict, other_controls=other_controls)

                    existingTableNames = tc["桌名"].values.astype(str).tolist()
                    if not tableName_allow_dup:
                        if tableName in existingTableNames:
                            print("桌名'{}'重复! 请重新输入! ".format(tableName))
                            confirm_tableName = False
                        else:
                            confirm_tableName = rtn_confirm_input(tableName)
                    else:
                        confirm_tableName = rtn_confirm_input(tableName)

                    
                
                confirm_tableMin = False
                while not confirm_tableMin:
                    tableMin = rtn_input_validation(rule="tableMin", title="最小载客量", rtn_constants_dict=rtn_constants_dict)
                    confirm_tableMin = rtn_confirm_input(tableMin)
                
                confirm_tableMax = False
                while not confirm_tableMax:
                    tableMax = rtn_input_validation(rule="tableMax", title="最大载客量", rtn_constants_dict=rtn_constants_dict, table_minimum_pax=tableMin)
                    confirm_tableMax = rtn_confirm_input(tableMax)
                
                tableId = str(uuid.uuid1().hex)
                table_time_created = pd.to_datetime(dt.datetime.now())

                table_sequencial_list = [tableId, tableName, tableMin, tableMax, table_time_created]

                new_table = {}
                for i in range(len(tc_columns)):
                    new_table.update({ tc_columns[i] : [table_sequencial_list[i]] })
                
                new_table = pd.DataFrame(new_table)
                prtdf(new_table.drop("桌子ID", axis=1))
                print()

                confirm_save, end_record_loop = rtn_confirm_save(tableName)

            else:
                if confirm_save:
                    table_concat = pd.concat([tc, new_table], ignore_index=True)
                    table_concat.sort_values(by="创建时间", ascending=False, inplace=True, ignore_index=True)
                    tc = table_concat
                    other_controls["tc"] = table_concat
                    print("'{}'桌台保存成功! ".format(tableName))
                else:
                    print("'{}'桌台未保存。".format(tableName))

        elif tc_select == 1:
            tableId = rtn_select_table(rtn_constants_dict=rtn_constants_dict, other_controls=other_controls)

            if isinstance(tableId, str):
                print()
                print()
                edit_options = option_num(["编辑其桌台的最小载客量", "编辑其桌台的最大载客量", "编辑桌名"])
                time.sleep(0.25)
                edit_select = option_limit(edit_options, input("在这里输入>>>: "))

                if edit_select == 0:
                    tableName = str(tc[tc["桌子ID"] == tableId]["桌名"].values[0])
                    old_tableMin = int(tc[tc["桌子ID"] == tableId]["最小载客量"].values[0])
                    print()
                    print("目前最小载客量: {}".format(old_tableMin))
                    print()
                    print()
                    end_record_loop = False
                    while not end_record_loop:
                        tableMin = rtn_input_validation(rule="tableMin", title="最小载客量", rtn_constants_dict=rtn_constants_dict)

                        print("确定把{}桌台的最小载客量从'{}'换成'{}'吗? ".format(tableName, old_tableMin, tableMin))

                        confirm_save, end_record_loop = rtn_confirm_save(tableMin)
                    else:
                        if confirm_save:
                            tableName = str(tc[tc["桌子ID"] == tableId]["桌名"].values[0])
                            tableMax = int(tc[tc["桌子ID"] == tableId]["最大载客量"].values[0])
                            table_timeCreated = pd.to_datetime(tc[tc["桌子ID"] == tableId]["创建时间"].values[0])

                            if tableMax < tableMin:
                                print("发现其桌台的最大载客量小于新最小载客量数值, 自动将最大载客量换成其最小载客量。")
                                tableMax = tableMin

                            sequence = [tableId, tableName, tableMin, tableMax, table_timeCreated]
                            modify = {}
                            for index in range(len(tc_columns)):
                                modify.update({ tc_columns[index] : [sequence[index]] })
                            
                            modify = pd.DataFrame(modify)
                            tc = other_controls["tc"].copy()
                            tc["桌子ID"] = tc["桌子ID"].astype(str)
                            tc = tc[tc["桌子ID"] != tableId]
                            tc = pd.concat([tc, modify], ignore_index=True)
                            tc["创建时间"] = pd.to_datetime(tc["创建时间"])
                            tc.sort_values(by="创建时间", ascending=False, inplace=True, ignore_index=True)
                            other_controls["tc"] = tc
                            print("编辑成功, 其桌台的最小载客量已从'{}'成功换成'{}'! ".format(old_tableMin, tableMin))
                        else:
                            print("编辑失败, 保留其原值。")
                
                elif edit_select == 1:
                    tableName = str(tc[tc["桌子ID"] == tableId]["桌名"].values[0])
                    old_tableMax = int(tc[tc["桌子ID"] == tableId]["最大载客量"].values[0])
                    print("目前最大载客量: {}".format(old_tableMax))
                    print()
                    print()
                    end_record_loop = False
                    while not end_record_loop:
                        tableMax = rtn_input_validation(rule="tableMax", title="最大载客量", rtn_constants_dict=rtn_constants_dict, table_minimum_pax=tableMin)

                        print("确定把{}桌台的最大载客量从'{}'换成'{}'吗? ".format(tableName, old_tableMax, tableMax))

                        confirm_save, end_record_loop = rtn_confirm_save(tableMax)
                    else:
                        if confirm_save:
                            tableName = str(tc[tc["桌子ID"] == tableId]["桌名"].values[0])
                            tableMin = int(tc[tc["桌子ID"] == tableId]["最小载客量"].values[0])
                            table_timeCreated = pd.to_datetime(tc[tc["桌子ID"] == tableId]["创建时间"].values[0])

                            if tableMax < tableMin:
                                print("发现其桌台的最小载客量大于新最大载客量数值,自动将最小载客量换成其最大载客量。")
                                tableMin = tableMax

                            sequence = [tableId, tableName, tableMin, tableMax, table_timeCreated]
                            modify = {}
                            for index in range(len(tc_columns)):
                                modify.update({ tc_columns[index] : [sequence[index]] })
                            
                            modify = pd.DataFrame(modify)
                            tc = other_controls["tc"].copy()
                            tc["桌子ID"] = tc["桌子ID"].astype(str)
                            tc = tc[tc["桌子ID"] != tableId]
                            tc = pd.concat([tc, modify], ignore_index=True)
                            tc["创建时间"] = pd.to_datetime(tc["创建时间"])
                            tc.sort_values(by="创建时间", ascending=False, inplace=True, ignore_index=True)
                            other_controls["tc"] = tc

                            print("编辑成功, 其桌台的最大载客量已从'{}'成功换成'{}'! ".format(old_tableMax, tableMax))
                        else:
                            print("编辑失败, 保留其原值。")

                elif edit_select == 2:
                    print()
                    print()
                    print("为了保持订单的同步性, 桌名不可更改。")
                    input("按回车键继续>>>: ")
                            
        elif tc_select == 2:
            print()
            print()
            print("为了保持订单的同步性, 创建的桌台不可删除。")
            input("按回车键继续>>>: ")

    else:
        return tc, other_controls

def rtn_edit_datetime(datetime=0):
    if isinstance(datetime, int):
        date = None
        while date == None:
            try:
                year = intRange(integer=input("请输入年份: "), lower=2000, upper=2037)
                month = intRange(integer=input("请输入月份: "), lower=1, upper=12)
                day = intRange(integer=input("请输入日: "), lower=1, upper=31)
                hour = intRange(integer=input("几点(24小时制)? : "), lower=0, upper=23)
                minute = intRange(integer=input("几分? : "), lower=0, upper=59)

                date = dt.datetime(year=int(year), month=int(month), day=int(day), hour=int(hour), minute=int(minute))
            
            except ValueError:
                print("日期输入有误, 请重新输入! ")
                date = None
            
        else:
            if not isinstance(date, int):
                date = pd.to_datetime(date)
            return date
    
    elif isinstance(datetime, pd.to_datetime):
        end_record_loop = False
        while not end_record_loop:
            modify_options = option_num(["修改时间'{}'".format(datetime.strftime("%Y-%m-%d %H:%M")), "不修改/退出修改时间"])
            time.sleep(0.25)
            modify_select = option_limit(modify_options, input("在这里输入>>>: "))

            if modify_select == 0:
                date = None
                while date == None:
                    try:
                        year = intRange(integer=input("请输入年份: "), lower=2000, upper=2037)
                        month = intRange(integer=input("请输入月份: "), lower=1, upper=12)
                        day = intRange(integer=input("请输入日: "), lower=1, upper=31)
                        hour = intRange(integer=input("几点(24小时制)? : "), lower=0, upper=23)
                        minute = intRange(integer=input("几分? : "), lower=0, upper=59)

                        date = dt.datetime(year=int(year), month=int(month), day=int(day), hour=int(hour), minute=int(minute))
                    
                    except ValueError:
                        print("日期输入有误, 请重新输入! ")
                        date = None
                    
                else:
                    date = pd.to_datetime(date)
                
                print("从'{}'改成'{}', 确定修改吗? ".format(datetime.strftime("%Y-%m-%d %H:%M"), date.strftime("%Y-%m-%d %H:%M")))
                confirm_save_new_datetime, end_record_loop = rtn_confirm_save(date.strftime("%Y-%m-%d %H:%M"))
            
            elif modify_select == 1:
                confirm_save_new_datetime = False
                end_record_loop = True
        else:
            if confirm_save_new_datetime:
                print("从'{}'改成'{}', 修改成功! ".format(datetime.strftime("%Y-%m-%d %H:%M"), date.strftime("%Y-%m-%d %H:%M")))
                datetime = date
            else:
                datetime = datetime
                print("修改不成功, 保留其原值。")
        
        return pd.to_datetime(datetime)

def rtn_fetch_database(google_auth, fernet_key, rtn_database_url, rtn_constants_dict):
    rtn_database_url = fernet_decrypt(rtn_database_url, fernet_key)
    rtn_db_sheetname = rtn_constants_dict["rtn_db_sheetname"]
    food_db_sheetname = rtn_constants_dict["food_db_sheetname"]
    sm_bd_sheetname = rtn_constants_dict["sm_bd_sheetname"]
    payment_sheetname = rtn_constants_dict["payment_sheetname"]

    try:
        rtn_database_sheet = google_auth.open_by_url(rtn_database_url)

        sheetnames = [
            rtn_db_sheetname,
            food_db_sheetname,
            sm_bd_sheetname,
            payment_sheetname
        ]

        sheetname_indexes = []
        for i in range(len(sheetnames)):
            sheetname_indexes += [rtn_database_sheet.worksheet(property="title", value=str(sheetnames[i])).index]
        
        db_keys = ["rtn_db", "food_db", "sm_bd", "payment"]

        database = {}
        for index in range(len(sheetnames)):
            database.update({ db_keys[index] : rtn_database_sheet[sheetname_indexes[index]].get_as_df() })

    except Exception as e:
        print()
        print("获取预订数据库失败, 错误描述如下: ")
        print()
        print(e)
        database = {}
    
    return database

def rtn_create_order(rtn_constants_dict, other_controls, google_auth, fernet_key, rtn_database_url):
    outletCode = str(rtn_constants_dict["outlet_code"]).strip().upper()
    orderAttributeList = str(rtn_constants_dict["order_attribute"]).strip().split(",")
    comingCnyEve = pd.to_datetime(str(rtn_constants_dict["coming_cny_eve_date"]))
    tc = other_controls["tc"]
    tc["最小载客量"] = tc["最小载客量"].astype(int)
    tc["最大载客量"] = tc["最大载客量"].astype(int)


    orderId = str(uuid.uuid1().hex)
    orderTimeCreated = pd.to_datetime(dt.datetime.now())

    confirm_orderAttribute = False
    while not confirm_orderAttribute:
        orderAttributeOptions = option_num(orderAttributeList)
        time.sleep(0.25)
        orderAttributeSelect = option_limit(orderAttributeOptions, input("在这里输入>>>: "))

        orderAttribute = orderAttributeList[orderAttributeSelect]

        confirm_orderAttribute = rtn_confirm_input(orderAttribute)

    print()
    confirm_reserveDatetime = False
    while not confirm_reserveDatetime:
        print("预订日期: ")
        reserveDatetime = rtn_edit_datetime()
        confirm_reserveDatetime = rtn_confirm_input(reserveDatetime.strftime("%Y-%m-%d %H:%M"))
    
    #预订除夕？
    CnyEveDate = dt.datetime(year=comingCnyEve.year, month=comingCnyEve.month, day=comingCnyEve.day)
    reserve_date = dt.datetime(year=reserveDatetime.year, month=reserveDatetime.month, day=reserveDatetime.day)

    if CnyEveDate == reserve_date:
        isCnyEve = 1
    else:
        isCnyEve = 0
    
    print()
    confirm_customerName = False
    while not confirm_customerName:
        customerName = rtn_input_validation(rule="searchString", title="客户名", space_removal=False, rtn_constants_dict=rtn_constants_dict)
        confirm_customerName = rtn_confirm_input(customerName)
    
    #订单号
    print()
    existingDb = rtn_fetch_database(google_auth=google_auth, fernet_key=fernet_key, rtn_database_url=rtn_database_url, rtn_constants_dict=rtn_constants_dict)["rtn_db"]
    existingDb = rtn_datetime_column_convert(existingDb)

    existingDb["预订除夕?"] = existingDb["预订除夕?"].astype(int)
    existingDb["订单属性"] = existingDb["订单属性"].astype(str)

    orderNumber_allow_dup = eval(str(rtn_constants_dict["order_number_allow_dup"]).strip().capitalize())

    confirm_orderNumber = False
    while not confirm_orderNumber:
        orderNumberOptions = option_num(["自动生成升序订单号", "自动生成随机订单号", "手动输入订单号"])
        time.sleep(0.25)
        orderNumberSelect = option_limit(orderNumberOptions, input("在这里输入>>>: "))

        if orderNumberSelect == 0:
            if not orderNumber_allow_dup:
                print("提示: 如果你删除了以往的订单, 删除的订单号可能会重新分配给新订单。")
                print()
                cnyEveDIStartNumber = int(rtn_constants_dict["cny_eve_dine_in_orderNumber_start"])
                cnyEveTAStartNumber = int(rtn_constants_dict["cny_eve_takeaway_orderNumber_start"])
                othersDIStartNumber = int(rtn_constants_dict["others_dine_in_orderNumber_start"])
                othersTAStartNumber = int(rtn_constants_dict["others_takeaway_orderNumber_start"])

                
                #订单属性, 预订除夕?
                if isCnyEve == 1:
                    if orderAttribute == orderAttributeList[0]:
                        #在除夕和堂食
                        filter = (existingDb["订单属性"] == orderAttributeList[0]) & (existingDb["预订除夕?"] == 1)
                        takeOrderNumberRule = cnyEveDIStartNumber
                    elif orderAttribute == orderAttributeList[1]:
                        #在除夕和打包
                        filter = (existingDb["订单属性"] == orderAttributeList[1]) & (existingDb["预订除夕?"] == 1)
                        takeOrderNumberRule = cnyEveTAStartNumber
                elif isCnyEve == 0:
                    if orderAttribute == orderAttributeList[0]:
                        #不在除夕和堂食
                        filter = (existingDb["订单属性"] == orderAttributeList[0]) & (existingDb["预订除夕?"] == 0)
                        takeOrderNumberRule = othersDIStartNumber
                    elif orderAttribute == orderAttributeList[1]:
                        #不在除夕和打包
                        filter = (existingDb["订单属性"] == orderAttributeList[1]) & (existingDb["预订除夕?"] == 0)
                        takeOrderNumberRule = othersTAStartNumber
                
                dbFiltered = existingDb.copy()
                dbFiltered = dbFiltered[filter]

                if dbFiltered.empty:
                    orderNumber = takeOrderNumberRule+1
                else:
                    orderNumber = takeOrderNumberRule+len(dbFiltered)+1
                
                confirm_orderNumber = True

            else:
                print("订单号允许重复,请选择别的订单号生成方案。")
                confirm_orderNumber = False
                
        elif orderNumberSelect == 1:
            orderNumber = str(uuid.uuid1().hex)
            confirm_orderNumber = rtn_confirm_input(orderNumber)
        
        else:
            orderNumber = rtn_input_validation(rule="orderNumber", title="订单号", space_removal=True, rtn_constants_dict=rtn_constants_dict)
            
            if not orderNumber_allow_dup:
                db_copy = existingDb.copy()
                db_copy["订单号"] = db_copy["订单号"].astype(str)
                if str(orderNumber) in db_copy["订单号"].values.tolist():
                    print("订单号重复! 请重新输入订单号! ")
                    confirm_orderNumber = False

                else:
                    confirm_orderNumber = rtn_confirm_input(orderNumber)
            else:
                confirm_orderNumber = rtn_confirm_input(orderNumber)

    else:
        print("此单订单号为: {}".format(orderNumber))
    
    print()
    confirm_customerPhone = False
    while not confirm_customerPhone:
        phoneOptionList = ["本地号码", "外国号码", "无预留号码"]
        phoneOptions = option_num(phoneOptionList)
        time.sleep(0.25)
        phoneOptionSelect = option_limit(phoneOptions, input("在这里输入>>>: "))

        if phoneOptionSelect == 0:
            customerPhone = rtn_input_validation(rule="localPhone", title="本地号码")
            confirm_customerPhone = rtn_confirm_input(customerPhone)
        
        elif phoneOptionSelect == 1:
            print("外国号码请添加国际代码,但无需'+'号。")
            customerPhone = rtn_input_validation(rule="searchString", title="外国号码", remove_spaces=True, rtn_constants_dict=rtn_constants_dict)
            if "+" in customerPhone:
                customerPhone = customerPhone.replace("+", "")
                print("'+'已自动移除, 无需添加'+'号。")
            
            confirm_customerPhone = rtn_confirm_input(customerPhone)
        
        else:
            print("无预留电话将以'-'呈现。")
            customerPhone = "-"
            confirm_customerPhone = True
    
    #本地电话?
    if phoneOptionList[phoneOptionSelect] == phoneOptionList[0]:
        isLocalPhone = 1
    else:
        isLocalPhone = 0
    
    #round
    print()
    round_time_param = str(rtn_constants_dict["cny_eve_round_time_param"]).split(",")
    
    round_time_dt_bind = {}
    for t in range(len(round_time_param)):
        dt_format = dt.datetime.strptime(round_time_param[t], "%Y-%m-%d %H:%M")
        round_sequence = "第{}轮".format(t+1)
        round_time_dt_bind.update({ dt_format : round_sequence  })

        if reserveDatetime in list(round_time_dt_bind.keys()):
            round = round_time_dt_bind[reserveDatetime]
        
        else:
            round = "无效"
    
    print("轮数选定为'{}'。".format(round))
    
    print()
    if orderAttribute == orderAttributeList[0]:
        confirm_adultPax = False
        while not confirm_adultPax:
            adultPax = rtn_input_validation(rule="pax", title="成人人数", other_controls=other_controls)
            confirm_adultPax = rtn_confirm_input(adultPax)
        
        confirm_childPax = False
        while not confirm_childPax:
            childPax = rtn_input_validation(rule="notAdultPax", title="儿童人数", other_controls=other_controls)
            confirm_childPax = rtn_confirm_input(childPax)
        
        confirm_toddlerPax = False
        while not confirm_toddlerPax:
            toddlerPax = rtn_input_validation(rule="notAdultPax", title="幼儿人数", other_controls=other_controls)
            confirm_toddlerPax = rtn_confirm_input(toddlerPax)
        
        confirm_totalPax = False
        while not confirm_totalPax:
            totalPax = rtn_input_validation(rule="pax", title="载客量", other_controls=other_controls)

            if totalPax < adultPax:
                print("载客量不能小于成人人数, 请重新输入! ")
            else:
                confirm_totalPax = rtn_confirm_input(totalPax)
    else:
        adultPax = 0
        childPax = 0
        toddlerPax = 0
        totalPax = 0

    #排桌位
    print()
    if orderAttribute == orderAttributeList[0]:
        db_copy = existingDb.copy()
        filter = (db_copy["预订时间"] == reserveDatetime)
        db_copy = db_copy[filter]
        db_copy["桌台"] = db_copy["桌台"].astype(str)

        confirm_table_assign = False
        while not confirm_table_assign:
            tableAssignOptions = option_num(["暂不排位", "排位"])
            time.sleep(0.25)
            tableAssignSelect = option_limit(tableAssignOptions, input("在这里输入>>>: "))

            if tableAssignSelect == 0:
                table_assign = "未排位"
                tableName = table_assign
                confirm_table_assign = True
            
            elif tableAssignSelect == 1:
                table_occupancy, unassigned_df = rtn_table_occupancy(rtn_db=db_copy, round=round, reserveTime=reserveDatetime, other_controls=other_controls)
                print("目前占位状况: ")
                prtdf(table_occupancy)
                print()
                print("现在开始选择你要排的台位...")
                print()
                table_assign = rtn_select_table(rtn_constants_dict=rtn_constants_dict, other_controls=other_controls)

                if not isinstance(table_assign, str):
                    print("选台位失败, 请重新选择! ")
                    confirm_table_assign = False
                
                else:
                    table_assigned = tc.copy()
                    table_assigned = table_assigned[table_assigned["桌子ID"] == str(table_assign)]
                    tableName = str(table_assigned["桌名"].values[0])
                    
                    if int(totalPax) <= int(table_assigned["最大载客量"].values[0]):
                        if int(totalPax) >= int(table_assigned["最小载客量"].values[0]):
                            if str(tableName) in db_copy["桌台"].values.astype(str).tolist():
                                print("排位失败, 该桌台在当天当轮已被占用! ")
                                print("请重新选择! ")
                                confirm_table_assign = False
                            else:
                                print("排位成功, 此订单已排位在桌台'{}'。".format(tableName))
                                confirm_table_assign = True
                        else:
                            print("排位失败, 该桌台太大了! ")
                            print("请重新选择! ")
                            confirm_table_assign = False
                    else:
                        print("排位失败, 该桌台太小了! ")
                        print("请重新选择! ")
                        confirm_table_assign = False
        else:
            if table_assign != "未排位":
                table_assign = str(tableName)
                print("该订单已排位在桌台'{}'。".format(tableName))
            else:
                print("该订单未排位。")

    else:
        table_assign = "无效"

    print()
    confirm_remark = False
    while not confirm_remark:
        remarkOption = option_num(["输入备注", "无备注"])
        time.sleep(0.25)
        remarkOptionSelect = option_limit(remarkOption, input("在这里输入>>>: "))

        if remarkOptionSelect == 0:
            remark = rtn_input_validation(rule="remark", title="备注", space_removal=False, rtn_constants_dict=rtn_constants_dict)

            if len(remark) == 0:
                print("无备注是吧? 自动用'-'呈现。")
                remark = "-"
            confirm_remark = rtn_confirm_input("该备注")
        
        else:
            print("无备注将用'-'呈现。")
            remark = "-"
            confirm_remark = True
    
    orderStatusList = str(rtn_constants_dict["order_status"]).strip().split(",")
    orderStatus = orderStatusList[0]

    #检查堂食超售率
    if orderAttribute == orderAttributeList[0]:
        overbooking_multiplier = float(rtn_constants_dict["dine_in_overbooking_multiplier"])
        db_copy = existingDb.copy()
        db_copy["预订时间"] = pd.to_datetime(db_copy["预订时间"])
        db_copy["订单属性"] = db_copy["订单属性"].astype(str)

        dine_in_filter = (db_copy["订单属性"] == orderAttributeList[0])
        db_copy = db_copy[dine_in_filter]

        similar_filter = (db_copy["预订时间"] == reserveDatetime) & (db_copy["载客量"] == totalPax)
        db_copy = db_copy[similar_filter]

        totalAvailTables = other_controls["tc"].copy()
        totalAvailTables = totalAvailTables[(totalAvailTables["最小载客量"] <= totalPax ) & (totalAvailTables["最大载客量"] >= totalPax)]
        totalAvailTables = len(totalAvailTables)

        totalBookingAllowable = int(np.floor(overbooking_multiplier*totalAvailTables))

        if len(db_copy)+1 > totalAvailTables:
            if len(db_copy)+1 > totalBookingAllowable:
                print("已经超出超售范围, 此订单无匹配桌台可供, 是否强制保存? ")
                saveRecordOptions = option_num(["强制保存", "不保存"])
                time.sleep(0.25)
                saveRecordSelect = option_limit(saveRecordOptions, input("在这里输入>>>: "))

                if saveRecordSelect == 0:
                    save_order = True
                else:
                    save_order = False
            else:
                print("已超出实际可供桌台范围, 但没有超出超售范围。")
                print("此订单可能无桌台可供。")
                print("因为没有超出超售范围, 因此可以被保存。")
                save_order = True
        else:
            print("超售率考量通过。")
            save_order = True
    
    else:
        save_order = True
    
    if save_order:
        sequencial_list = [outletCode, 
                           orderId, 
                           orderTimeCreated, 
                           orderAttribute, 
                           reserveDatetime, 
                           isCnyEve, 
                           customerName, 
                           orderNumber, 
                           customerPhone,
                           isLocalPhone,
                           round,
                           adultPax,
                           childPax,
                           toddlerPax,
                           totalPax,
                           table_assign,
                           remark,
                           orderStatus
                          ]
        
        order_columns = str(rtn_constants_dict["order_columns"]).strip().split(",")

        new_order = {}
        for i in range(len(order_columns)):
            new_order.update({ order_columns[i] : [sequencial_list[i]] })

        new_order = pd.DataFrame(new_order)

        order_concat = pd.concat([existingDb, new_order], ignore_index=True)
        order_concat.sort_values(by="预订时间", ascending=False, inplace=True, ignore_index=True)

        print("订单ID: {}保存成功! ".format(orderId))
    else:
        print("订单ID: {}保存不成功, 已舍弃。".format(orderId))
        order_concat = existingDb

    return order_concat

def rtn_select_order(order_concat, rtn_constants_dict):
    print()
    order_concat = rtn_datetime_column_convert(order_concat)
    order_copy = order_concat.copy()
    order_copy = rtn_datetime_column_convert(order_copy)

    if order_concat.empty:
        print("无订单。")
        choose_order_input = None
    
    else:
        options = option_num(["按订单号搜索", "按电话搜索", "按轮数搜索", "按订单状态搜索", "按备注内容搜索", "按照规则重新排列所有订单搜索", "按姓名搜索", "退出订单选取"])
        time.sleep(0.25)
        userInput = option_limit(options, input("在这里输入>>>: "))

        if userInput == 0:
            order_copy["订单号"] = order_copy["订单号"].astype(str)
            orderNumber = rtn_input_validation(rule="orderNumber", title="订单号", space_removal=True, rtn_constants_dict=rtn_constants_dict)
            order_copy = order_copy[order_copy["订单号"].str.contains(str(orderNumber))]
        
        elif userInput == 1:
            order_copy["电话"] = order_copy["电话"].astype(str)
            phone = rtn_input_validation(rule="searchString", title="电话号码", space_removal=True, rtn_constants_dict=rtn_constants_dict)
            order_copy = order_copy[order_copy["电话"].str.contains(str(phone))]

        elif userInput == 2:
            order_copy["轮数"] = order_copy["轮数"].astype(str)

            uniqueChoices = np.unique(order_copy["轮数"].values)
            uniqueChoicesNum = option_num(uniqueChoices.astype(str).tolist())
            time.sleep(0.25)
            roundSelect = option_limit(uniqueChoicesNum, input("在这里输入>>>: "))

            order_copy = order_copy[order_copy["轮数"] == uniqueChoices[roundSelect]]
        
        elif userInput == 3:
            order_copy["订单状态"] = order_copy["订单状态"].astype(str)
            uniqueChoices = np.unique(order_copy["订单状态"].values)
            uniqueChoicesNum = option_num(uniqueChoices.astype(str).tolist())
            time.sleep(0.25)
            statusSelect = option_limit(uniqueChoicesNum, input("在这里输入>>>: "))

            order_copy = order_copy[order_copy["订单状态"] == uniqueChoices[statusSelect]]

        elif userInput == 4:
            order_copy["备注"] = order_copy["备注"].astype(str)
            text_search = rtn_input_validation(rule="searchString", title="一些备注内容", space_removal=True, rtn_constants_dict=rtn_constants_dict)
            order_copy = order_copy[order_copy["备注"].str.contains(text_search)]
        
        elif userInput == 5:
            filterOptions = option_num([
                "按订单创建时间从小到大排列",
                "按订单创建时间从大到小排列",
                "按订单属性从大到小排列",
                "按订单属性从小到大排列",
                "按姓名从小到大排列",
                "按姓名从大到小排列",
                "按订单号从小到大排列",
                "按订单号从大到小排列",
                "按载客量从小到大排列",
                "按载客量从大到小排列",
            ])

            time.sleep(0.25)
            filter_select = option_limit(filterOptions, input("在这里输入>>>: "))

            if filter_select == 0:
                order_copy.sort_values(by="订单创建时间", ascending=True, inplace=True, ignore_index=True)
            elif filter_select == 1:
                order_copy.sort_values(by="订单创建时间", ascending=False, inplace=True, ignore_index=True)
            elif filter_select == 2:
                order_copy.sort_values(by="订单属性", ascending=True, inplace=True, ignore_index=True)
            elif filter_select == 3:
                order_copy.sort_values(by="订单属性", ascending=False, inplace=True, ignore_index=True)
            elif filter_select == 4:
                order_copy.sort_values(by="姓名", ascending=True, inplace=True, ignore_index=True)
            elif filter_select == 5:
                order_copy.sort_values(by="姓名", ascending=False, inplace=True, ignore_index=True)
            elif filter_select == 6:
                order_copy.sort_values(by="订单号", ascending=True, inplace=True, ignore_index=True)
            elif filter_select == 7:
                order_copy.sort_values(by="订单号", ascending=False, inplace=True, ignore_index=True)
            elif filter_select == 8:
                order_copy.sort_values(by="载客量", ascending=True, inplace=True, ignore_index=True)
            elif filter_select == 9:
                order_copy.sort_values(by="载客量", ascending=False, inplace=True, ignore_index=True)

        elif userInput == 6:
            order_copy["姓名"] = order_copy["姓名"].astype(str)
            customerName = rtn_input_validation(rule="searchString", title="客户名", space_removal=False, rtn_constants_dict=rtn_constants_dict)
            order_copy = order_copy[order_copy["姓名"].str.contains(customerName)]

        elif userInput == 7:
            print("好的, 已退出订单选取。")
            choose_order_input = None
        
        if userInput != 7:
            if order_copy.empty:
                print("未找到任何匹配的订单。")
                choose_order_input = None
            else:
                order_copy.reset_index(inplace=True)
                order_copy.drop("index", axis=1, inplace=True)

                order_list = []
                for index in range(len(order_copy)):
                    order_list += ["'订单号: {}, {}({}人)预订在{}, 电话: {}'".format(order_copy.iloc[index, 7], order_copy.iloc[index, 6], order_copy.iloc[index, 14], order_copy.iloc[index, 4].strftime("%Y-%m-%d %H:%M"), order_copy.iloc[index, 8])]
                
                choose_order_options = option_num(order_list)
                time.sleep(0.25)
                choose_order_input = option_limit(choose_order_options, input("在这里输入>>>: "))
        else:
            choose_order_input = None

    if isinstance(choose_order_input, int):
        return str(order_copy.iloc[choose_order_input, 1])
    else:
        return None

def rtn_order_filter(google_auth, fernet_key, rtn_database_url, rtn_constants_dict, other_controls, orderId):
    order_columns = str(rtn_constants_dict["order_columns"]).split(",")
    order_columns = str(order_columns).replace("桌台", "桌台名")
    order_columns = ast.literal_eval(order_columns)

    acm = other_controls["acm"].copy()
    sm = other_controls["sm"].copy()

    acm["菜品ID"] = acm["菜品ID"].astype(str)
    sm["菜品ID"] = sm["菜品ID"].astype(str)

    tc = other_controls["tc"].copy()
    tc["桌子ID"] = tc["桌子ID"].astype(str)
    tc["桌名"] = tc["桌名"].astype(str)

    db = rtn_fetch_database(google_auth=google_auth, fernet_key=fernet_key, rtn_database_url=rtn_database_url, rtn_constants_dict=rtn_constants_dict)
    rtn_db = rtn_datetime_column_convert(db["rtn_db"].copy())
    rtn_db["订单ID"] = rtn_db["订单ID"].astype(str)

    food_db = db["food_db"].copy()
    food_db["订单ID"] = food_db["订单ID"].astype(str)
    food_db["点餐INDEX"] = food_db["点餐INDEX"].astype(int)
    food_db["点餐ID"] = food_db["点餐ID"].astype(str)
    food_db["菜品ID"] = food_db["菜品ID"].astype(str)

    sm_bd = db["sm_bd"].copy()
    sm_bd["订单ID"] = sm_bd["订单ID"].astype(str)
    sm_bd["点餐ID"] = sm_bd["点餐ID"].astype(str)

    payment = rtn_datetime_column_convert(db["payment"].copy())
    payment["订单ID"] = payment["订单ID"].astype(str)

    if isinstance(orderId, str):
        rtn_db = rtn_db[rtn_db["订单ID"] == orderId]
        food_db = food_db[food_db["订单ID"] == orderId]

        food_db = food_db[food_db["订单ID"] == orderId]
        food_db.sort_values(by="点餐INDEX", ascending=True, ignore_index=True, inplace=True)

        sm_bd = sm_bd[sm_bd["订单ID"] == orderId]

        payment = payment[payment["订单ID"] == orderId]

    else:
        rtn_db = None
        food_db = None
        sm_bd = None
        payment = None

    return rtn_db, food_db, sm_bd, payment    

def rtn_point_zero_five_round(x):
    return normal_round(normal_round(x / 0.05) * 0.05, -int(np.floor(np.log10(0.05))))

def rtn_food_order_parser(food_db, sm_bd, payment, other_controls):
    sm = other_controls["sm"].copy()
    sm["菜品ID"] = sm["菜品ID"].astype(str)

    acm = other_controls["acm"].copy()
    acm["菜品ID"] = acm["菜品ID"].astype(str)

    if food_db.empty:
        display_food_db = -404

    else:
        display_food_db = food_db.copy()

        foodNames = []
        if not food_db.empty:
            for i in range(len(food_db)):
                if int(food_db.iloc[i, 4]) == 0:
                    foodNames += [str(acm[acm["菜品ID"] == str(food_db.iloc[i, 3])]["菜名"].values[0])]
                else:
                    foodNames += [str(sm[sm["菜品ID"] == str(food_db.iloc[i, 3])]["套餐名"].values[0])]
            
        display_food_db["菜名/套餐名"] = foodNames

        display_food_db = display_food_db[["点餐ID", "订单ID", "菜品ID", "菜名/套餐名", "点餐INDEX", "套餐?", "换菜?", "菜品属性", "数量", "±价", "折扣", "税前价格", "服务费", "GST", "税后价格", "备注"]]

    if not sm_bd.empty:
        display_sm_bd = sm_bd.copy()
        
        sm_bd_df = pd.DataFrame(columns=["点餐ID", "套餐名", "菜肴Index", "菜肴名", "备注"])
        for i in range(len(display_sm_bd)):
            foodOrderId = str(display_sm_bd.iloc[i, 1])
            foodId = str(food_db[food_db["点餐ID"] == str(foodOrderId)]["菜品ID"].values[0])
            foodName = str(sm[sm["菜品ID"] == str(foodId)]["套餐名"].values[0])
            foodQty = float(food_db[food_db["点餐ID"] == str(foodOrderId)]["数量"].values[0])

            foodItems = str(display_sm_bd.iloc[i, 2]).split(",")
            foodRemarks = str(display_sm_bd.iloc[i, 3]).split(",")

            sm_bd_reconstruct = {}
            for item in range(len(foodItems)):
                itemName = foodItems[item]
                remark = foodRemarks[item]
                foodIndex = item+1

                sm_bd_reconstruct.update({ 
                    "点餐ID" : [foodOrderId],
                    "数量" : [foodQty],
                    "套餐名" : [foodName],
                    "菜肴Index" : [foodIndex],
                    "菜肴名" : [itemName],
                    "备注": [remark]
                    })
                
                sm_bd_reconstruct = pd.DataFrame(sm_bd_reconstruct)

                sm_bd_df = pd.concat([sm_bd_df, sm_bd_reconstruct], ignore_index=True)
        
        sm_bd_df.sort_values(by=["点餐ID", "菜肴Index"], ascending=True, inplace=True, ignore_index=True)

    else:
        sm_bd_df = -404
    
    if isinstance(display_food_db, pd.DataFrame):
        display_food_db["税后价格"] = display_food_db["税后价格"].astype(float)
        display_food_db["服务费"] = display_food_db["服务费"].astype(float)
        display_food_db["GST"] = display_food_db["GST"].astype(float)
        display_food_db["税前价格"] = display_food_db["税前价格"].astype(float)

        payment["金额"] = payment["金额"].astype(float)
        total_payment_required = rtn_point_zero_five_round(float(format(float(display_food_db["税后价格"].sum()), ".4f")))
        already_paid = float(format(float(payment["金额"].sum()), ".2f"))

        if already_paid == total_payment_required:
            pending_payment = 0

        else:
            pending_payment = total_payment_required - already_paid
            pending_payment = float(format(float(pending_payment), ".4f"))

        total_svc = float(format(float(display_food_db["服务费"].sum()), ".4f"))
        total_gst = float(format(float(display_food_db["GST"].sum()), ".4f"))
        total_subtotal = float(format(float(display_food_db["税前价格"].sum()), ".4f"))

        payment_info = {
            "总税前价格" : total_subtotal,
            "总服务费" : total_svc,
            "总GST" : total_gst,
            "总税后价格" : total_payment_required,
            "已付金额" : already_paid,
            "需付金额" : pending_payment,
        }
    else:
        already_paid = float(format(float(payment["金额"].sum()), ".2f"))
        payment_info = {
            "总税前价格" : 0.00,
            "总服务费" : 0.00,
            "总GST" : 0.00,
            "总税后价格" : 0.00,
            "已付金额" : already_paid,
            "需付金额" : 0.00,
        }
    
    return display_food_db, food_db, sm_bd_df, sm_bd, payment_info

def rtn_edit_food_change_remark(rtn_constants_dict, food_items, foodItemRemark=""):
    print()
    print("正在编辑套餐内菜肴备注。")
    print()
    foodItem = food_items.split(",")

    if foodItemRemark == "":
        foodRemark = []
        for _ in range(len(foodItem)):
            foodRemark += [""]
    else:
        foodRemark = foodItemRemark.split(",")
    
    options = []
    for item in range(len(foodItem)):
        options += ["编辑{}".format(foodItem[item])]
    
    options += ["我已完成编辑菜肴备注"]

    userSelect = 0
    while userSelect != len(foodItem):
        for index in range(len(foodItem)):
            print("{} : {}".format(foodItem[index], foodRemark[index]))

        option = option_num(options)
        time.sleep(0.25)
        userSelect  = option_limit(option, input("在这里输入>>>: "))

        if userSelect != len(foodItem):
            confirm_remark = False
            while not confirm_remark:
                remark = rtn_input_validation(rule="remark", title="备注", space_removal=False, rtn_constants_dict=rtn_constants_dict)
                print("你确定吗? ")
                print("你确定{}的备注内容是'{}'吗?".format(foodItem[userSelect], remark))
                
                confirm_remark = rtn_confirm_input("该备注")
            else:
                foodRemark[userSelect] = remark
    else:
        return_foodRemark = ""

        for index in range(len(foodRemark)):
            if index != len(foodRemark)-1:
                return_foodRemark += "{},".format(foodRemark[index])
            else:
                return_foodRemark += "{}".format(foodRemark[index])
        
        return return_foodRemark

def rtn_food_change_handler(foodName, rtn_constants_dict, fi):
    userInput = 0
    while userInput != 3:
        try:
            food_items = food_items
        except:
            food_items = fi

        
        if food_items == -404:
            food_items_list = fi.split(",")
        else:
            food_items_list = food_items.split(",")
        
        try:
            food_remark = food_remark
            food_remark_list = food_remark.split(",")
        except:
            food_remark_list = np.repeat("", len(food_items_list)).tolist()
        
        for index in range(len(food_items_list)):
            print("{} : {}".format(food_items_list[index], food_remark_list[index]))
        
        print()
        print()
        options = option_num(["编辑其菜肴", "编辑其备注", "我已完成编辑菜肴和备注"])
        time.sleep(0.25)
        userInput = option_limit(options, input("在这里输入>>>: "))
        
        if userInput == 0:
            try:
                food_items = food_items
            except:
                food_items = fi
            
            if food_items == -404:
                food_items = fi
            else:
                food_items = food_items

            food_items = rtn_edit_food_item(foodName=foodName, rtn_constants_dict=rtn_constants_dict, food_items=food_items)

        elif userInput == 1:
            try:
                food_remark = food_remark
            except:
                food_remark = ""

            try:
                food_items = food_items
            except:
                food_items = fi
            
            if food_items == -404:
                food_items = fi
            else:
                food_items = food_items

            food_remark = rtn_edit_food_change_remark(rtn_constants_dict=rtn_constants_dict, food_items=food_items, foodItemRemark=food_remark)

        elif userInput == 2:
            try:
                food_items = food_items
            except:
                food_items = -404
            
            if isinstance(food_items, int):
                if food_items == -404:
                    print("你选择了有换菜, 请编辑其菜肴后再选择此项。")
                
            else:
                userInput = 3
    else:
        try:
            food_remark = food_remark
        except:
            food_remark = -404

        if isinstance(food_remark, list):
            return_foodRemark = ""

            for index in range(len(food_items)):
                if index != len(food_items)-1:
                    return_foodRemark += "{},".format(food_remark[index])
                else:
                    return_foodRemark += "{}".format(food_remark[index])

            food_remark = return_foodRemark
        
        else:
            if isinstance(food_remark, int):
                return_foodRemark = ""

                for index in range(len(food_items_list)):
                    if index != len(food_items_list)-1:
                        return_foodRemark += ","
                    else:
                        return_foodRemark += ""
                
                food_remark = return_foodRemark
            else:
                pass

        return food_items, food_remark

def rtn_select_foodOrderId(display_food_db):
    if isinstance(display_food_db, pd.DataFrame):
        print()
        print("目前点餐详情: ")
        prtdf(display_food_db)

        display_food_db.sort_values(by="点餐INDEX", ascending=True, inplace=True, ignore_index=True)
        options = []
        for i in range(len(display_food_db)):
            foodIndex = display_food_db.iloc[i, 4]
            foodName = display_food_db.iloc[i, 3]
            options += ["选择第{}样菜'{}'".format(foodIndex, foodName)]
        
        options += ["退出菜品选取"]

        option = option_num(options)
        time.sleep(0.25)
        userSelect = option_limit(option, input("在这里输入>>>: "))

        if userSelect != len(option)-1:
            foodOrderId = str(display_food_db.iloc[userSelect, 0])
        
        else:
            print("好的,已退出选择菜品。")
            foodOrderId = None
    else:
        print("目前未点餐。")
        foodOrderId = None

    return foodOrderId

def rtn_edit_food_order(google_auth, fernet_key, rtn_database_url, order_concat, rtn_constants_dict, other_controls):
    acm = other_controls["acm"].copy()
    acm["菜品ID"] = acm["菜品ID"].astype(str)

    sm = other_controls["sm"].copy()
    sm["菜品ID"] = sm["菜品ID"].astype(str)

    svc_rate = float(rtn_constants_dict["service_charge_multiplier"])
    gst_rate = float(rtn_constants_dict["gst_multiplier"])

    food_db_columns = str(rtn_constants_dict["food_db_columns"]).split(",")
    sm_bd_columns = str(rtn_constants_dict["sm_bd_columns"]).split(",")

    fo_select = 0
    while fo_select != 1:
        try:
            slice_food_db = slice_food_db
        except:
            slice_food_db = -404
        
        try:
            slice_sm_bd = slice_sm_bd
        except:
            slice_sm_bd = -404
        
        if isinstance(slice_food_db, pd.DataFrame):
            option = ["保存并退出订餐操作"]
        else:
            option = ["查看订单", "退出订餐操作"]

        first_options = option_num(option)
        time.sleep(0.25)
        fo_select = option_limit(first_options, input("在这里输入>>>: "))

        if len(option) == 2:
            if fo_select == 0:
                orderId = rtn_select_order(order_concat=order_concat, rtn_constants_dict=rtn_constants_dict)

                if isinstance(orderId, str):
                    rtn_db, food_db, sm_bd, payment = rtn_order_filter(google_auth=google_auth, fernet_key=fernet_key, rtn_database_url=rtn_database_url, rtn_constants_dict=rtn_constants_dict, other_controls=other_controls, orderId=orderId)

                    try:
                        slice_food_db = slice_food_db
                    except:
                        slice_food_db = food_db
                    
                    try:
                        slice_sm_bd = slice_sm_bd
                    except:
                        slice_sm_bd = sm_bd

                    if not isinstance(slice_food_db, pd.DataFrame):
                        slice_food_db = food_db
                    
                    if not isinstance(slice_sm_bd, pd.DataFrame):
                        slice_sm_bd = sm_bd

                    display_food_db, food_db, sm_bd_df, sm_bd, payment_info = rtn_food_order_parser(food_db=slice_food_db, sm_bd=slice_sm_bd, payment=payment, other_controls=other_controls)
                    
                    if isinstance(display_food_db, pd.DataFrame):
                        display_food_db["点餐ID"] = display_food_db["点餐ID"].astype(str)
                        display_food_db["菜品ID"] = display_food_db["菜品ID"].astype(str)

                    second_select = 0
                    while second_select != 2:
                        print()
                        print()
                        print("订单详情: ")
                        for index in range(len(rtn_db)):
                            for column in range(len(rtn_db.columns)):
                                print("{}: {}".format(rtn_db.columns[column], rtn_db.iloc[index, column]))
                        
                        print()
                        print("系统备注: 「预订除夕?」和「本地号码?」列, 1表是, 0表不是。")
                        print()
                        
                        if isinstance(display_food_db, pd.DataFrame):
                            print()
                            print("目前点餐详情: ")
                            prtdf(display_food_db.drop(["点餐ID", "订单ID", "菜品ID"], axis=1))
                            print("系统备注: 「套餐?」和「换菜?」列, 1表是, 0表不是。")
                        else:
                            print()
                            print("目前未点餐。")
                        
                        if isinstance(sm_bd_df, pd.DataFrame):
                            print()
                            print("套餐内菜肴详情: ")
                            prtdf(sm_bd_df.drop(["点餐ID"], axis=1))
                        else:
                            print()
                        
                        print()
                        print("财务详情: ")
                        for key, value in payment_info.items():
                            print("{} : {}".format(key, format(float(value), ".2f")))
                            
                        print()

                        second_options = option_num(["加餐", "编辑餐品", "返回上一菜单"])
                        time.sleep(0.25)
                        second_select = option_limit(second_options, input("在这里输入>>>: "))

                        if second_select == 0:
                            end_record_loop = False
                            while not end_record_loop:
                                third_option = option_num(["单点", "套餐"])
                                time.sleep(0.25)
                                third_select = option_limit(third_option, input("在这里输入>>>: "))

                                if third_select == 0:
                                    foodId = rtn_select_food(ala_carte=True, rtn_constants_dict=rtn_constants_dict, other_controls=other_controls)
                                    isSet = 0
                                    hasFoodChange = 0
                                else:
                                    isSet = 1
                                    confirm_food_change = False
                                    while not confirm_food_change:
                                        print("首先，有换菜吗? ")
                                        food_change_option = ["有换菜", "无换菜"]
                                        food_change_options = option_num(food_change_option)
                                        time.sleep(0.25)
                                        food_change_select = option_limit(food_change_options, input("在这里输入>>>: "))
                                        
                                        if food_change_select == 0:
                                            hasFoodChange = 1
                                        else:
                                            hasFoodChange = 0
                                        
                                        print("你确定吗? ")
                                        confirm_food_change = rtn_confirm_input(food_change_option[food_change_select])

                                    foodId = rtn_select_food(ala_carte=False, rtn_constants_dict=rtn_constants_dict, other_controls=other_controls)

                                if isinstance(foodId, str):
                                    foodOrderId = str(uuid.uuid1().hex)

                                    if third_select == 1:
                                        if hasFoodChange == 1:
                                            originalFoodItems = str(sm[sm["菜品ID"] == foodId]["菜肴"].values[0])
                                            foodName = str(sm[sm["菜品ID"] == foodId]["套餐名"].values[0])

                                            food_item, food_remark = rtn_food_change_handler(foodName=foodName, rtn_constants_dict=rtn_constants_dict, fi=originalFoodItems)
                                        
                                        else:
                                            food_item = str(sm[sm["菜品ID"] == foodId]["菜肴"].values[0])
                                            foodName = str(sm[sm["菜品ID"] == foodId]["套餐名"].values[0])

                                            confirm_foodItemRemark = False
                                            while not confirm_foodItemRemark:
                                                food_remark = rtn_edit_food_change_remark(rtn_constants_dict=rtn_constants_dict, food_items=food_item, foodItemRemark="")
                                                print("你确定吗? ")
                                                confirm_foodItemRemark = rtn_confirm_input("这些备注")

                                    if isSet == 0:
                                        base_price = float(format(float(acm[acm["菜品ID"] == str(foodId)]["价格"].values[0]), ".2f"))
                                    
                                    else:
                                        base_price = float(format(float(sm[sm["菜品ID"] == str(foodId)]["价格"].values[0]), ".2f"))

                                    orderAttribute = rtn_db["订单属性"].values[0]

                                    if orderAttribute == "堂食":
                                        confirm_foodOrderAttribute = False
                                        while not confirm_foodOrderAttribute:
                                            foodOrderAttr_option = ["堂食(有服务费)", "打包(无服务费)"]
                                            foodOrderAttr_options = option_num(foodOrderAttr_option)
                                            time.sleep(0.25)
                                            foodOrderAttr_select = option_limit(foodOrderAttr_options, input("在这里输入>>>: "))

                                            if foodOrderAttr_select == 0:
                                                foodOrderAtrribute = "堂食"
                                            else:
                                                foodOrderAtrribute = "打包"

                                            print("你确定吗? ")
                                            confirm_foodOrderAttribute = rtn_confirm_input(foodOrderAttr_option[foodOrderAttr_select])
                                        
                                        else:
                                            print("菜品属性为'{}'。".format(foodOrderAtrribute))
                                    else:
                                        foodOrderAtrribute = "打包"
                                    
                                    confirm_food_add_on = False
                                    while not confirm_food_add_on:
                                        foodAddOnOption = ["有加减价", "无加减价"]
                                        foodAddOnOptions = option_num(foodAddOnOption)
                                        time.sleep(0.25)
                                        foodAddOnSelect = option_limit(foodAddOnOptions, input("在这里输入>>>: "))

                                        if foodAddOnSelect == 0:
                                            hasFoodAddOn = True
                                        else:
                                            hasFoodAddOn = False

                                        if hasFoodAddOn:
                                            foodAddOn = rtn_input_validation(rule="float", title="价格(加价正数,减价负数)")
                                            foodAddOn = float(format(foodAddOn, ".2f"))

                                            if foodAddOn < -1*base_price:
                                                print("减到倒贴钱了, 请重新输入! ")
                                                confirm_food_add_on = False
                                            else:
                                                print("你确定吗? ")
                                                print("目前基础价格: {}".format(base_price))
                                                print("你输入的加减价格: {}".format(foodAddOn))
                                                print("影响后的基础价格: {}".format(format(base_price+foodAddOn, ".2f")))
                                                print()
                                                confirm_food_add_on = rtn_confirm_input(foodAddOn)
                                        else:
                                            foodAddOn = 0
                                            print("你确定吗? ")
                                            confirm_food_add_on = rtn_confirm_input("没有加减价")
                                    else:
                                        print("加减价为: {}".format(foodAddOn))

                                    confirm_foodQty = False
                                    while not confirm_foodQty:
                                        foodQty = rtn_input_validation(rule="foodQty", title="数量", rtn_constants_dict=rtn_constants_dict)

                                        print("确定这个数量吗? ")
                                        confirm_foodQty = rtn_confirm_input(foodQty)
                                    
                                    confirm_food_discount = False
                                    while not confirm_food_discount:
                                        food_discount_options = option_num(["有折扣", "无折扣"])
                                        time.sleep(0.25)
                                        food_discount_select = option_limit(food_discount_options, input("在这里输入>>>: "))

                                        if food_discount_select == 0:
                                            discount_option = option_num(["按价格给予折扣", "按百分比给予折扣"])
                                            time.sleep(0.25)
                                            discount_option_select = option_limit(discount_option, input("在这里输入>>>: "))

                                            if discount_option_select == 0:
                                                food_discount = rtn_input_validation(rule="float", title="折扣价格")
                                                food_discount = abs(float(format(food_discount, ".2f")))

                                                if food_discount > ((base_price+foodAddOn)*foodQty):
                                                    print("折扣到倒贴钱了, 请重新输入! ")
                                                    confirm_food_discount = False
                                                else:
                                                    print("你确定吗? ")
                                                    print("目前价格: {}".format(format(((base_price+foodAddOn)*foodQty), ".2f")))
                                                    print("你输入的折扣价格: {}".format(food_discount))
                                                    print("受影响后的价格: {}".format(format(((base_price+foodAddOn)*foodQty)-food_discount, ".2f")))
                                                    print()
                                                    confirm_food_discount = rtn_confirm_input(food_discount)
                                            else:
                                                food_discount_percentage = intRange(integer=input("请输入折扣百分比: "), lower=0, upper=100)
                                                food_discount_percentage = food_discount_percentage/100
                                                food_discount = float(format(((base_price+foodAddOn)*foodQty)*food_discount_percentage, ".2f"))

                                                print("你确定吗? ")
                                                print("目前价格: {}".format(format(((base_price+foodAddOn)*foodQty), ".2f")))
                                                print("你输入的折扣百分比: {}%, 即折扣{}".format(food_discount_percentage*100, food_discount))
                                                print("受影响后的价格: {}".format(format(((base_price+foodAddOn)*foodQty)-food_discount, ".2f")))
                                                print()
                                                confirm_food_discount = rtn_confirm_input("{}%".format(food_discount_percentage*100))
                                        else:
                                            food_discount = 0
                                            print("你确定吗? ")
                                            confirm_food_discount = rtn_confirm_input("无折扣")
                                    else:
                                        print("折扣为: {}".format(food_discount))
                                    
                                    total_subtotal = (base_price+foodAddOn)*foodQty-food_discount
                                    total_subtotal = float(format(total_subtotal, ".4f"))

                                    if foodOrderAtrribute == "堂食":
                                        svc_charge = svc_rate - 1
                                        total_svc = float(format(total_subtotal*svc_charge, ".4f"))

                                    else:
                                        total_svc = 0.00
                                    
                                    gst_charge = gst_rate - 1
                                    total_gst = float(format((total_subtotal+total_svc)*gst_charge, ".4f"))

                                    total_payment = rtn_point_zero_five_round(float(format(total_subtotal+total_svc+total_gst, ".4f")))

                                    confirm_remark = False
                                    while not confirm_remark:
                                        remarkOption = option_num(["输入备注", "无备注"])
                                        time.sleep(0.25)
                                        remarkOptionSelect = option_limit(remarkOption, input("在这里输入>>>: "))

                                        if remarkOptionSelect == 0:
                                            remark = rtn_input_validation(rule="remark", title="备注", space_removal=False, rtn_constants_dict=rtn_constants_dict)
                                            if len(remark) == 0:
                                                print("无备注是吧? 自动用'-'呈现。")
                                                remark = "-"

                                            confirm_remark = rtn_confirm_input("该备注")
                                        
                                        else:
                                            print("无备注将用'-'呈现。")
                                            remark = "-"
                                            confirm_remark = True
                                    
                                    foodOrderIndex = len(food_db)+1

                                    food_db_sequence = [
                                        foodOrderId,
                                        orderId,
                                        foodOrderIndex,
                                        foodId,
                                        isSet,
                                        hasFoodChange,
                                        foodOrderAtrribute,
                                        foodQty,
                                        foodAddOn,
                                        food_discount,
                                        total_subtotal,
                                        total_svc,
                                        total_gst,
                                        total_payment,
                                        remark]

                                    new_food_db = {}
                                    for index in range(len(food_db_columns)):
                                        new_food_db.update({ food_db_columns[index] : [food_db_sequence[index]] })
                                    
                                    new_food_db = pd.DataFrame(new_food_db)
                                    
                                    food_db_concat = pd.concat([food_db, new_food_db], ignore_index=True)
                                    food_db_concat.sort_values(by="点餐INDEX", ascending=True, inplace=True, ignore_index=True)

                                    if isSet:
                                        sm_bd_sequence = [
                                            orderId,
                                            foodOrderId,
                                            food_item,
                                            food_remark
                                        ]

                                        new_sm_bd = {}  
                                        for index in range(len(sm_bd_columns)):
                                            new_sm_bd.update({ sm_bd_columns[index] : [sm_bd_sequence[index]] })
                                        
                                        new_sm_bd = pd.DataFrame(new_sm_bd)

                                        sm_bd_concat = pd.concat([sm_bd, new_sm_bd], ignore_index=True)
                                        sm_bd_concat.sort_values(by="点餐ID", ascending=True, inplace=True, ignore_index=True)
                                    
                                    try:
                                        sm_bd_concat = sm_bd_concat
                                    except:
                                        sm_bd_concat = sm_bd
                                    
                                    unconf_display_food_db, food_db_concat, unconf_sm_bd_df, sm_bd_concat, unconf_payment_info = rtn_food_order_parser(food_db=food_db_concat, sm_bd=sm_bd_concat, payment=payment, other_controls=other_controls)
                                    
                                    print()
                                    print()
                                    print()
                                    print("订单详情: ")
                                    for index in range(len(rtn_db)):
                                        for column in range(len(rtn_db.columns)):
                                            print("{}: {}".format(rtn_db.columns[column], rtn_db.iloc[index, column]))

                                    print()
                                    print("系统备注: 「预订除夕?」和「本地号码?」列, 1表是, 0表不是。")
                                    print()
                                    print()
                                    if isinstance(unconf_display_food_db, pd.DataFrame):
                                        print()
                                        print("更新后点餐详情: ")
                                        prtdf(unconf_display_food_db.drop(["点餐ID", "订单ID", "菜品ID"], axis=1))
                                        print("系统备注: 「套餐?」和「换菜?」列, 1表是, 0表不是。")
                                        print()
                                    else:
                                        print()

                                    if isinstance(unconf_sm_bd_df, pd.DataFrame):
                                        print()
                                        print("更新后套餐内菜肴详情: ")
                                        prtdf(unconf_sm_bd_df.drop(["点餐ID"], axis=1))
                                    else:
                                        print()
                                    print()
                                    print("更新后的财务详情: ")
                                    for key, value in unconf_payment_info.items():
                                        print("{} : {}".format(key, format(float(value), ".2f")))
                                    
                                    print()
                                    print()
                                    confirm_save, end_record_loop = rtn_confirm_save("此记录")
                                else:
                                    confirm_save = False
                                    end_record_loop = True
                            else:
                                if confirm_save:
                                    food_db = food_db_concat
                                    sm_bd = sm_bd_concat
                                    display_food_db = unconf_display_food_db
                                    sm_bd_df = unconf_sm_bd_df
                                    payment_info = unconf_payment_info

                                    display_food_db["点餐ID"] = display_food_db["点餐ID"].astype(str)
                                
                                else:
                                    print("记录未保存。")
        
                        elif second_select == 1:
                            print("选择要编辑的菜品。")
                            foodOrderId = rtn_select_foodOrderId(display_food_db=display_food_db)

                            if isinstance(foodOrderId, str):
                                fourth_select = 0
                                while fourth_select != 5:

                                    try:
                                        if display_food_db[display_food_db["点餐ID"] == foodOrderId].empty:
                                            fourth_select = 5
                                        
                                        else:
                                            prtdf(display_food_db[display_food_db["点餐ID"] == foodOrderId])

                                            foodId = str(display_food_db[display_food_db["点餐ID"] == foodOrderId]["菜品ID"].values[0])
                                            
                                            fourth_options = option_num(["换个菜品", "删除菜品", "修改数量/加减价/折扣", "修改菜品属性", "编辑备注", "返回上一菜单"])
                                            time.sleep(0.25)
                                            fourth_select = option_limit(fourth_options, input("在这里输入>>>: "))

                                            if fourth_select == 0:
                                                end_record_loop = False
                                                while not end_record_loop:
                                                    fifth_option = option_num(["单点", "套餐"])
                                                    time.sleep(0.25)
                                                    fifth_select = option_limit(fifth_option, input("在这里输入>>>: "))

                                                    if fifth_select == 0:
                                                        new_foodId = rtn_select_food(ala_carte=True, rtn_constants_dict=rtn_constants_dict, other_controls=other_controls)
                                                        isSet = 0
                                                        hasFoodChange = 0

                                                    else:
                                                        isSet = 1
                                                        confirm_food_change = False
                                                        while not confirm_food_change:
                                                            print("首先，有换菜吗? ")
                                                            food_change_option = ["有换菜", "无换菜"]
                                                            food_change_options = option_num(food_change_option)
                                                            time.sleep(0.25)
                                                            food_change_select = option_limit(food_change_options, input("在这里输入>>>: "))
                                                            
                                                            if food_change_select == 0:
                                                                hasFoodChange = 1
                                                            else:
                                                                hasFoodChange = 0
                                                            
                                                            print("你确定吗? ")
                                                            confirm_food_change = rtn_confirm_input(food_change_option[food_change_select])
                                                        
                                                        new_foodId = rtn_select_food(ala_carte=False, rtn_constants_dict=rtn_constants_dict, other_controls=other_controls)

                                                    if acm[acm["菜品ID"] == foodId].empty:
                                                        old_foodName = str(sm[sm["菜品ID"] == foodId]["套餐名"].values[0])
                                                    else:
                                                        old_foodName = str(acm[acm["菜品ID"] == foodId]["菜名"])
                                                        
                                                    if isSet == 1:
                                                        new_foodName = str(sm[sm["菜品ID"] == new_foodId]["套餐名"].values[0])
                                                    else:
                                                        new_foodName = str(acm[acm["菜品ID"] == new_foodId]["菜名"].values[0])
                                                
                                                    if hasFoodChange == 1:
                                                        originalFoodItems = str(sm[sm["菜品ID"] == new_foodId]["菜肴"].values[0])
                                                        foodName = str(sm[sm["菜品ID"] == new_foodId]["套餐名"].values[0])

                                                        food_item, food_remark = rtn_food_change_handler(foodName=foodName, rtn_constants_dict=rtn_constants_dict, fi=originalFoodItems)
                                                    
                                                    else:
                                                        food_item = str(sm[sm["菜品ID"] == new_foodId]["菜肴"].values[0])
                                                        foodName = str(sm[sm["菜品ID"] == new_foodId]["套餐名"].values[0])

                                                        confirm_foodItemRemark = False
                                                        while not confirm_foodItemRemark:
                                                            food_remark = rtn_edit_food_change_remark(rtn_constants_dict=rtn_constants_dict, food_items=food_item, foodItemRemark="")
                                                            print("你确定吗? ")
                                                            confirm_foodItemRemark = rtn_confirm_input("这些备注")
                                                
                                                    if isSet == 0:
                                                        base_price = float(acm[acm["菜品ID"] == str(new_foodId)]["价格"].values[0])
                                                        base_price = format(base_price, ".2f")
                                                        base_price = float(base_price)
                                                
                                                    else:
                                                        base_price = float(sm[sm["菜品ID"] == str(new_foodId)]["价格"].values[0])
                                                        base_price = format(base_price, ".2f")
                                                        base_price = float(base_price)
                                                
                                                    foodOrderAtrribute = str(display_food_db[display_food_db["点餐ID"] == foodOrderId]["菜品属性"].values[0])
                                                    foodAddOn = float(display_food_db[display_food_db["点餐ID"] == foodOrderId]["±价"].values[0])
                                                    foodAddOn = format(foodAddOn, ".2f")
                                                    foodAddOn = float(foodAddOn)

                                                    if foodAddOn < -1*base_price:
                                                        print("换了菜品之后, 加减价减到倒贴钱了。")
                                                        print("请重新修改加减价! ")
                                                        confirm_food_add_on = False
                                                        while not confirm_food_add_on:
                                                            foodAddOn = rtn_input_validation(rule="float", title="价格(加价正数,减价负数)")
                                                            foodAddOn = float(format(foodAddOn, ".2f"))

                                                            if foodAddOn < -1*base_price:
                                                                print("减到倒贴钱了, 请重新输入! ")
                                                                confirm_food_add_on = False
                                                            else:
                                                                print("你确定吗? ")
                                                                print("目前基础价格: {}".format(base_price))
                                                                print("你输入的加减价格: {}".format(foodAddOn))
                                                                print("影响后的基础价格: {}".format(format(base_price+foodAddOn, ".2f")))
                                                                print()
                                                                confirm_food_add_on = rtn_confirm_input(foodAddOn)
                                                    
                                                    foodQty = int(display_food_db[display_food_db["点餐ID"] == foodOrderId]["数量"].values[0])
                                                    food_discount = float(display_food_db[display_food_db["点餐ID"] == foodOrderId]["折扣"].values[0])

                                                    if food_discount > ((base_price+foodAddOn)*foodQty):
                                                        print("换了菜品之后, 折扣到倒贴钱了。")
                                                        print("请重新输入! ")
                                                        confirm_food_discount = False
                                                        while not confirm_food_discount:
                                                            discount_option = option_num(["按价格给予折扣", "按百分比给予折扣"])
                                                            time.sleep(0.25)
                                                            discount_option_select = option_limit(discount_option)

                                                            if discount_option_select == 0:
                                                                food_discount = rtn_input_validation(rule="float", title="折扣价格")
                                                                food_discount = abs(float(format(food_discount, ".2f")))

                                                                if food_discount > ((base_price+foodAddOn)*foodQty):
                                                                    print("折扣到倒贴钱了, 请重新输入! ")
                                                                    confirm_food_discount = False
                                                                else:
                                                                    print("你确定吗? ")
                                                                    print("目前价格: {}".format(format(((base_price+foodAddOn)*foodQty), ".2f")))
                                                                    print("你输入的折扣价格: {}".format(food_discount))
                                                                    print("受影响后的价格: {}".format(format(((base_price+foodAddOn)*foodQty)-food_discount, ".2f")))
                                                                    print()
                                                                    confirm_food_discount = rtn_confirm_input(food_discount)
                                                            else:
                                                                food_discount_percentage = intRange(integer=input("请输入折扣百分比: "), lower=0, upper=100)
                                                                food_discount_percentage = food_discount_percentage/100
                                                                food_discount = float(format(((base_price+foodAddOn)*foodQty)*food_discount_percentage, ".2f"))

                                                                print("你确定吗? ")
                                                                print("目前价格: {}".format(format(((base_price+foodAddOn)*foodQty), ".2f")))
                                                                print("你输入的折扣百分比: {}%, 即折扣{}".format(food_discount_percentage*100, food_discount))
                                                                print("受影响后的价格: {}".format(format(((base_price+foodAddOn)*foodQty)-food_discount, ".2f")))
                                                                print()
                                                                confirm_food_discount = rtn_confirm_input("{}%".format(food_discount_percentage*100))

                                                    total_subtotal = (base_price+foodAddOn)*foodQty-food_discount  
                                                    total_subtotal = float(format(total_subtotal, ".4f"))

                                                    if foodOrderAtrribute == "堂食":
                                                        svc_charge = svc_rate - 1
                                                        total_svc = float(format(total_subtotal*svc_charge, ".4f"))

                                                    else:
                                                        total_svc = 0.00  

                                                    gst_charge = gst_rate - 1
                                                    total_gst = float(format((total_subtotal+total_svc)*gst_charge, ".4f"))

                                                    total_payment = rtn_point_zero_five_round(float(format(total_subtotal+total_svc+total_gst, ".4f")))

                                                    remark = str(display_food_db[display_food_db["点餐ID"] == foodOrderId]["备注"].values[0])
                                                    foodOrderIndex = int(display_food_db[display_food_db["点餐ID"] == foodOrderId]["点餐INDEX"].values[0])
                                                    
                                                    print()
                                                    print("修改详情: ")
                                                    print("点餐ID: {}".format(foodOrderId))
                                                    print("订单ID : {}".format(orderId))
                                                    print("点餐INDEX: {}".format(foodOrderIndex))
                                                    print("菜品ID: {}".format(new_foodId))
                                                    print("菜名/套餐名: {}".format(new_foodName))
                                                    print("套餐?: {}".format(isSet))
                                                    print("换菜?: {}".format(hasFoodChange))
                                                    print("菜品属性: {}".format(foodOrderAtrribute))
                                                    print("数量: {}".format(foodQty))
                                                    print("±价: {}".format(foodAddOn))
                                                    print("折扣: {}".format(food_discount))
                                                    print("税前价格: {}".format(total_subtotal))
                                                    print("服务费: {}".format(total_svc))
                                                    print("GST: {}".format(total_gst))
                                                    print("税后价格: {}".format(total_payment))
                                                    print("备注: {}".format(remark))
                                                    print()
                                                    
                                                    if isSet:
                                                        print("套餐内菜肴修改详情: ")
                                                        food_items = food_item.split(",")
                                                        food_remarks = food_remark.split(",")
                                                        for i in range(len(food_items)):
                                                            print("{} : {}".format(food_items[i], food_remarks[i]))
                                                    
                                                    print("你确定把'{}'改成'{}'吗? ".format(old_foodName, new_foodName))
                                                    confirm_save, end_record_loop = rtn_confirm_save("此修改")
                                                else:
                                                    if confirm_save:
                                                        food_db_sequence = [
                                                            foodOrderId,
                                                            orderId,
                                                            foodOrderIndex,
                                                            new_foodId,
                                                            isSet,
                                                            hasFoodChange,
                                                            foodOrderAtrribute,
                                                            foodQty,
                                                            foodAddOn,
                                                            food_discount,
                                                            total_subtotal,
                                                            total_svc,
                                                            total_gst,
                                                            total_payment,
                                                            remark]

                                                        new_food_db = {}
                                                        for index in range(len(food_db_columns)):
                                                            new_food_db.update({ food_db_columns[index] : [food_db_sequence[index]] })
                                                        
                                                        new_food_db = pd.DataFrame(new_food_db)

                                                        food_db = food_db[food_db["点餐ID"] != foodOrderId]
                                                        food_db_concat = pd.concat([food_db, new_food_db], ignore_index=True)
                                                        food_db_concat.sort_values(by="点餐INDEX", ascending=True, inplace=True, ignore_index=True)
                                                        food_db = food_db_concat

                                                        if isSet == 1:
                                                            sm_bd_sequence = [
                                                                orderId,
                                                                foodOrderId,
                                                                food_item,
                                                                food_remark
                                                            ]

                                                            new_sm_bd = {}
                                                            for index in range(len(sm_bd_columns)):
                                                                new_sm_bd.update({ sm_bd_columns[index] : [sm_bd_sequence[index]] })
                                                            
                                                            new_sm_bd = pd.DataFrame(new_sm_bd)
                                                            sm_bd = sm_bd[sm_bd["点餐ID"] != foodOrderId]
                                                            sm_bd_concat = pd.concat([sm_bd, new_sm_bd], ignore_index=True)
                                                            sm_bd_concat.sort_values(by="点餐ID", ascending=True, inplace=True, ignore_index=True)
                                                            sm_bd = sm_bd_concat
                                                        
                                                        display_food_db, food_db, sm_bd_df, sm_bd, payment_info = rtn_food_order_parser(food_db=food_db, sm_bd=sm_bd, payment=payment, other_controls=other_controls)

                                                        print("修改成功, '{}'已成功改成'{}'".format(old_foodName, new_foodName))
                                                    else:
                                                        print("修改失败, 保留'{}'".format(old_foodName))

                                            elif fourth_select == 1:
                                                foodName = str(display_food_db[display_food_db["菜品ID"] == foodId]["菜名/套餐名"].values[0])
                                                isSet = int(display_food_db[display_food_db["点餐ID"] == foodOrderId]["套餐?"].values[0])

                                                print("确定删除'{}'吗? ".format(foodName))
                                                confirm_del = rtn_confirm_input(foodName)
                                                
                                                if confirm_del:
                                                    food_db = food_db[food_db["点餐ID"] != foodOrderId]
                                                    food_db["点餐INDEX"] = food_db["点餐INDEX"].astype(int)
                                                    food_db.sort_values(by="点餐INDEX", ascending=True, inplace=True, ignore_index=True)
                                                    food_db["点餐INDEX"] = np.arange(1, len(food_db)+1)
                                                    food_db["点餐INDEX"] = food_db["点餐INDEX"].astype(int)

                                                    if isSet == 1:
                                                        sm_bd = sm_bd[sm_bd["点餐ID"] != foodOrderId]
                                                    
                                                    display_food_db, food_db, sm_bd_df, sm_bd, payment_info = rtn_food_order_parser(food_db=food_db, sm_bd=sm_bd, payment=payment, other_controls=other_controls)
                                                    print("'{}'删除成功! ".format(foodName))
                                                
                                                else:
                                                    print("'{}'未删除! ".format(foodName))

                                            elif fourth_select == 2:
                                                end_record_loop = False
                                                while not end_record_loop:
                                                    foodName = str(display_food_db[display_food_db["点餐ID"] == foodOrderId]["菜名/套餐名"].values[0])
                                                    foodOrderAttribute = str(display_food_db[display_food_db["点餐ID"] == foodOrderId]["菜品属性"].values[0])
                                                    isSet = int(display_food_db[display_food_db["点餐ID"] == foodOrderId]["套餐?"].values[0])
                                                    foodOrderIndex = int(display_food_db[display_food_db["点餐ID"] == foodOrderId]["点餐INDEX"].values[0])
                                                    remark = str(display_food_db[display_food_db["点餐ID"] == foodOrderId]["备注"].values[0])
                                                    hasFoodChange = int(display_food_db[display_food_db["点餐ID"] == foodOrderId]["换菜?"].values[0])

                                                    if isSet == 0:
                                                        base_price = float(format(float(acm[acm["菜品ID"] == str(foodId)]["价格"].values[0]), ".2f"))
                                                    
                                                    else:
                                                        base_price = float(format(float(sm[sm["菜品ID"] == str(foodId)]["价格"].values[0]), ".2f"))


                                                    confirm_food_add_on = False
                                                    while not confirm_food_add_on:
                                                        foodAddOnOption = ["有加减价", "无加减价"]
                                                        foodAddOnOptions = option_num(foodAddOnOption)
                                                        time.sleep(0.25)
                                                        foodAddOnSelect = option_limit(foodAddOnOptions, input("在这里输入>>>: "))

                                                        if foodAddOnSelect == 0:
                                                            hasFoodAddOn = True
                                                        else:
                                                            hasFoodAddOn = False

                                                        if hasFoodAddOn:
                                                            foodAddOn = rtn_input_validation(rule="float", title="价格(加价正数,减价负数)")
                                                            foodAddOn = float(format(foodAddOn, ".2f"))

                                                            if foodAddOn < -1*base_price:
                                                                print("减到倒贴钱了, 请重新输入! ")
                                                                confirm_food_add_on = False
                                                            else:
                                                                print("你确定吗? ")
                                                                print("目前基础价格: {}".format(base_price))
                                                                print("你输入的加减价格: {}".format(foodAddOn))
                                                                print("影响后的基础价格: {}".format(format(base_price+foodAddOn, ".2f")))
                                                                print()
                                                                confirm_food_add_on = rtn_confirm_input(foodAddOn)
                                                        else:
                                                            foodAddOn = 0
                                                            print("你确定吗? ")
                                                            confirm_food_add_on = rtn_confirm_input("没有加减价")
                                                    else:
                                                        print("加减价为: {}".format(foodAddOn))


                                                    confirm_foodQty = False
                                                    while not confirm_foodQty:
                                                        foodQty = rtn_input_validation(rule="foodQty", title="数量", rtn_constants_dict=rtn_constants_dict)

                                                        print("确定这个数量吗? ")
                                                        confirm_foodQty = rtn_confirm_input(foodQty)
                                                    
                                                    confirm_food_discount = False

                                                    while not confirm_food_discount:
                                                        food_discount_options = option_num(["有折扣", "无折扣"])
                                                        time.sleep(0.25)
                                                        food_discount_select = option_limit(food_discount_options, input("在这里输入>>>: "))

                                                        if food_discount_select == 0:
                                                            discount_option = option_num(["按价格给予折扣", "按百分比给予折扣"])
                                                            time.sleep(0.25)
                                                            discount_option_select = option_limit(discount_option)

                                                            if discount_option_select == 0:
                                                                food_discount = rtn_input_validation(rule="float", title="折扣价格")
                                                                food_discount = abs(float(format(food_discount, ".2f")))

                                                                if food_discount > ((base_price+foodAddOn)*foodQty):
                                                                    print("折扣到倒贴钱了, 请重新输入! ")
                                                                    confirm_food_discount = False
                                                                else:
                                                                    print("你确定吗? ")
                                                                    print("目前价格: {}".format(format(((base_price+foodAddOn)*foodQty), ".2f")))
                                                                    print("你输入的折扣价格: {}".format(food_discount))
                                                                    print("受影响后的价格: {}".format(format(((base_price+foodAddOn)*foodQty)-food_discount, ".2f")))
                                                                    print()
                                                                    confirm_food_discount = rtn_confirm_input(food_discount)
                                                            else:
                                                                food_discount_percentage = intRange(integer=input("请输入折扣百分比: "), lower=0, upper=100)
                                                                food_discount_percentage = food_discount_percentage/100
                                                                food_discount = float(format(((base_price+foodAddOn)*foodQty)*food_discount_percentage, ".2f"))

                                                                print("你确定吗? ")
                                                                print("目前价格: {}".format(format(((base_price+foodAddOn)*foodQty), ".2f")))
                                                                print("你输入的折扣百分比: {}%, 即折扣{}".format(food_discount_percentage*100, food_discount))
                                                                print("受影响后的价格: {}".format(format(((base_price+foodAddOn)*foodQty)-food_discount, ".2f")))
                                                                print()
                                                                confirm_food_discount = rtn_confirm_input("{}%".format(food_discount_percentage*100))
                                                        else:
                                                            food_discount = 0
                                                            print("你确定吗? ")
                                                            confirm_food_discount = rtn_confirm_input("无折扣")
                                                    else:
                                                        print("折扣为: {}".format(food_discount))
                                                    
                                                    total_subtotal = (base_price+foodAddOn)*foodQty-food_discount
                                                    total_subtotal = float(format(total_subtotal, ".2f"))

                                                    if foodOrderAttribute == "堂食":
                                                        svc_charge = svc_rate - 1
                                                        total_svc = float(format(total_subtotal*svc_charge, ".4f"))

                                                    else:
                                                        total_svc = 0.00
                                                    
                                                    gst_charge = gst_rate - 1
                                                    total_gst = float(format((total_subtotal+total_svc)*gst_charge, ".4f"))

                                                    total_payment = rtn_point_zero_five_round(float(format(total_subtotal+total_svc+total_gst, ".4f")))

                                                    print("修改详情: ")
                                                    print("菜名/套餐名: {}".format(foodName))
                                                    print("菜品属性: {}".format(foodOrderAttribute))
                                                    print("数量: {}".format(foodQty))
                                                    print("±价: {}".format(foodAddOn))
                                                    print("折扣: {}".format(food_discount))
                                                    print("税前价格: {}".format(total_subtotal))
                                                    print("服务费: {}".format(total_svc))
                                                    print("GST: {}".format(total_gst))
                                                    print("税后价格: {}".format(total_payment))

                                                    print("你确定这些修改吗? ")
                                                    confirm_save, end_record_loop = rtn_confirm_save("这些修改")
                                                else:
                                                    if confirm_save:
                                                        food_db_sequence = [
                                                            foodOrderId,
                                                            orderId,
                                                            foodOrderIndex,
                                                            foodId,
                                                            isSet,
                                                            hasFoodChange,
                                                            foodOrderAttribute,
                                                            foodQty,
                                                            foodAddOn,
                                                            food_discount,
                                                            total_subtotal,
                                                            total_svc,
                                                            total_gst,
                                                            total_payment,
                                                            remark]
                                                        
                                                        modify_food_db = {}
                                                        for index in range(len(food_db_columns)):
                                                            modify_food_db.update({ food_db_columns[index] : [ food_db_sequence[index] ] })

                                                        modify_food_db = pd.DataFrame(modify_food_db)
                                                        food_db = food_db[food_db["点餐ID"] != foodOrderId]
                                                        food_db_concat = pd.concat([food_db, modify_food_db], ignore_index=True)
                                                        food_db_concat.sort_values(by="点餐INDEX", ascending=True, inplace=True, ignore_index=True)

                                                        food_db = food_db_concat
                                                    
                                                        display_food_db, food_db, sm_bd_df, sm_bd, payment_info = rtn_food_order_parser(food_db=food_db, sm_bd=sm_bd, payment=payment, other_controls=other_controls)

                                                        print("修改成功! ")
                                                    else:
                                                        print("未修改! ")
                                            
                                            elif fourth_select == 3:
                                                orderAttribute = str(rtn_db[rtn_db["订单ID"] == orderId]["订单属性"].values[0])
                                                original_foodOrderAttribute = str(display_food_db[display_food_db["点餐ID"] == foodOrderId]["菜品属性"].values[0])

                                                confirm_foodOrderAttribute = False
                                                while not confirm_foodOrderAttribute:
                                                    if orderAttribute == "打包":
                                                        if original_foodOrderAttribute == "打包":
                                                            print("订单属性是'打包', 菜品属性必须是'打包'")
                                                            input("按回车键继续>>>: ")
                                                            confirm_foodOrderAttribute = True
                                                    else:
                                                        option = option_num(["打包(无服务费)", "堂食(有服务费)"])
                                                        time.sleep(0.25)
                                                        foaSelect = option_limit(option, input("在这里输入>>>: "))

                                                        if foaSelect == 0:
                                                            foodOrderAttribute = "打包"
                                                        else:
                                                            foodOrderAttribute = "堂食"
                                                        
                                                        print("你确定吗? ")
                                                        confirm_foodOrderAttribute = rtn_confirm_input(foodOrderAttribute)
                                                
                                                foodName = str(display_food_db[display_food_db["点餐ID"] == foodOrderId]["菜名/套餐名"].values[0])
                                                isSet = int(display_food_db[display_food_db["点餐ID"] == foodOrderId]["套餐?"].values[0])
                                                foodOrderIndex = int(display_food_db[display_food_db["点餐ID"] == foodOrderId]["点餐INDEX"].values[0])
                                                remark = str(display_food_db[display_food_db["点餐ID"] == foodOrderId]["备注"].values[0])
                                                hasFoodChange = int(display_food_db[display_food_db["点餐ID"] == foodOrderId]["换菜?"].values[0])

                                                foodAddOn = float(display_food_db[display_food_db["点餐ID"] == foodOrderId]["±价"].values[0])
                                                food_discount = float(display_food_db[display_food_db["点餐ID"] == foodOrderId]["折扣"].values[0])

                                                if isSet == 0:
                                                        base_price = float(format(float(acm[acm["菜品ID"] == str(foodId)]["价格"].values[0]), ".2f"))
                                                    
                                                else:
                                                    base_price = float(format(float(sm[sm["菜品ID"] == str(foodId)]["价格"].values[0]), ".2f"))

                                                total_subtotal = (base_price+foodAddOn)*foodQty-food_discount
                                                total_subtotal = float(format(total_subtotal, ".4f"))

                                                if foodOrderAttribute == "堂食":
                                                    svc_charge = svc_rate - 1
                                                    total_svc = float(format(total_subtotal*svc_charge, ".4f"))

                                                else:
                                                    total_svc = 0.00
                                                
                                                gst_charge = gst_rate - 1
                                                total_gst = float(format((total_subtotal+total_svc)*gst_charge, ".4f"))

                                                total_payment = rtn_point_zero_five_round(float(format(total_subtotal+total_svc+total_gst), ".4f"))
                                            
                                                print("修改详情: ")
                                                print("菜名/套餐名: {}".format(foodName))
                                                print("菜品属性: {}".format(foodOrderAttribute))
                                                print("数量: {}".format(foodQty))
                                                print("±价: {}".format(foodAddOn))
                                                print("折扣: {}".format(food_discount))
                                                print("税前价格: {}".format(total_subtotal))
                                                print("服务费: {}".format(total_svc))
                                                print("GST: {}".format(total_gst))
                                                print("税后价格: {}".format(total_payment))

                                                food_db_sequence = [
                                                            foodOrderId,
                                                            orderId,
                                                            foodOrderIndex,
                                                            foodId,
                                                            isSet,
                                                            hasFoodChange,
                                                            foodOrderAttribute,
                                                            foodQty,
                                                            foodAddOn,
                                                            food_discount,
                                                            total_subtotal,
                                                            total_svc,
                                                            total_gst,
                                                            total_payment,
                                                            remark]

                                                modify_food_db = {}
                                                for index in range(len(food_db_columns)):
                                                    modify_food_db.update({ food_db_columns[index] : [ food_db_sequence[index] ] })

                                                modify_food_db = pd.DataFrame(modify_food_db)
                                                food_db = food_db[food_db["点餐ID"] != foodOrderId]
                                                food_db_concat = pd.concat([food_db, modify_food_db], ignore_index=True)
                                                food_db_concat.sort_values(by="点餐INDEX", ascending=True, inplace=True, ignore_index=True)

                                                display_food_db, food_db, sm_bd_df, sm_bd, payment_info = rtn_food_order_parser(food_db=food_db, sm_bd=sm_bd, payment=payment, other_controls=other_controls)

                                            elif fourth_select == 4:
                                                isSet = int(display_food_db[display_food_db["点餐ID"] == foodOrderId]["套餐?"].values[0])

                                                if isSet == 1:
                                                    option = ["改该套餐内菜肴备注", "改该菜品备注", "完成并退出修改备注"]
                                                else:
                                                    option = ["改该菜品备注", "完成并退出修改备注"]
                                                
                                                sixth_select = 0
                                                while sixth_select != len(option)-1:
                                                    options = option_num(option)
                                                    time.sleep(0.25)
                                                    sixth_select = option_limit(options, input("在这里输入>>>: "))

                                                    if len(option) == 3:
                                                        if sixth_select == 0:
                                                            food_item = str(sm_bd[sm_bd["点餐ID"] == foodOrderId]["菜肴"].values[0])
                                                            food_remark = str(sm_bd[sm_bd["点餐ID"] == foodOrderId]["备注"].values[0])

                                                            to_save_food_remark = rtn_edit_food_change_remark(rtn_constants_dict=rtn_constants_dict, food_items=food_item, foodItemRemark=food_remark)

                                                        elif sixth_select == 1:
                                                            print("当前菜品备注: ")
                                                            current_remark = str(display_food_db[display_food_db["点餐ID"] == foodOrderId]["备注"].values[0])
                                                            print(current_remark)
                                                            print()
                                                            to_save_remark = rtn_input_validation(rule="remark", title="备注", space_removal=False, rtn_constants_dict=rtn_constants_dict)
                                                    else:
                                                        if sixth_select == 0:
                                                            print("当前菜品备注: ")
                                                            current_remark = str(display_food_db[display_food_db["点餐ID"] == foodOrderId]["备注"].values[0])
                                                            print(current_remark)
                                                            print()
                                                            to_save_remark = rtn_input_validation(rule="remark", title="备注", space_removal=False, rtn_constants_dict=rtn_constants_dict)
                                                else:
                                                    try:
                                                        to_save_food_remark = to_save_food_remark
                                                        saving_food_remark = True
                                                    except:
                                                        saving_food_remark = False
                                                    
                                                    try:
                                                        to_save_remark = to_save_remark
                                                        saving_remark = True
                                                    except:
                                                        saving_remark = False
                                                    
                                                    if saving_food_remark:
                                                        sequence = [
                                                            orderId,
                                                            foodOrderId,
                                                            food_item,
                                                            to_save_food_remark]
                                                        
                                                        new_sm_bd = {}
                                                        for index in range(len(sm_bd_columns)):
                                                            new_sm_bd.update({ sm_bd_columns[index] : [ sequence[index] ] })
                                                        
                                                        new_sm_bd = pd.DataFrame(new_sm_bd)
                                                        sm_bd = sm_bd[sm_bd["点餐ID"] != foodOrderId]
                                                        sm_bd_concat = pd.concat([sm_bd, new_sm_bd], ignore_index=True)
                                                        sm_bd_concat.sort_values(by="点餐ID", ascending=True, inplace=True, ignore_index=True)
                                                        sm_bd = sm_bd_concat
                                                    
                                                    if saving_remark:
                                                        foodOrderIndex = int(display_food_db[display_food_db["点餐ID"] == foodOrderId]["点餐INDEX"].values[0])
                                                        isSet = int(display_food_db[display_food_db["点餐ID"] == foodOrderId]["套餐?"].values[0])
                                                        hasFoodChange = int(display_food_db[display_food_db["点餐ID"] == foodOrderId]["换菜?"].values[0])
                                                        foodOrderAttribute = str(display_food_db[display_food_db["点餐ID"] == foodOrderId]["菜品属性"].values[0])
                                                        foodQty = int(display_food_db[display_food_db["点餐ID"] == foodOrderId]["数量"].values[0])
                                                        foodAddOn = float(display_food_db[display_food_db["点餐ID"] == foodOrderId]["±价"].values[0])
                                                        food_discount = float(display_food_db[display_food_db["点餐ID"] == foodOrderId]["折扣"].values[0])
                                                        total_subtotal = float(display_food_db[display_food_db["点餐ID"] == foodOrderId]["税前价格"].values[0])
                                                        total_svc = float(display_food_db[display_food_db["点餐ID"] == foodOrderId]["服务费"].values[0])
                                                        total_gst = float(display_food_db[display_food_db["点餐ID"] == foodOrderId]["GST"].values[0])
                                                        total_payment = float(display_food_db[display_food_db["点餐ID"] == foodOrderId]["税后价格"].values[0])
                                                        
                                                        sequence = [
                                                            foodOrderId,
                                                            orderId,
                                                            foodOrderIndex,
                                                            foodId,
                                                            isSet,
                                                            hasFoodChange,
                                                            foodOrderAttribute,
                                                            foodQty,
                                                            foodAddOn,
                                                            food_discount,
                                                            total_subtotal,
                                                            total_svc,
                                                            total_gst,
                                                            total_payment,
                                                            to_save_remark]

                                                        modify_food_db = {}
                                                        for index in range(len(food_db_columns)):
                                                            modify_food_db.update({ food_db_columns[index] : [ sequence[index]] })
                                                        
                                                        modify_food_db = pd.DataFrame(modify_food_db)
                                                        food_db = food_db[food_db["点餐ID"] != foodOrderId]
                                                        food_db_concat = pd.concat([food_db, modify_food_db], ignore_index=True)
                                                        food_db_concat.sort_values(by="点餐INDEX", ascending=True, inplace=True, ignore_index=True)
                                                        food_db = food_db_concat
                                                    
                                                    display_food_db, food_db, sm_bd_df, sm_bd, payment_info = rtn_food_order_parser(food_db=food_db, sm_bd=sm_bd, payment=payment, other_controls=other_controls)
                                                    print("备注类操作完成。")
                                    except Exception as e:
                                        print()
                                        print()
                                        print()
                                        print("错误描述: ")
                                        print(e)
                                        print()
                                        print()
                                        fourth_select = 5

                    else:
                        slice_food_db = food_db
                        slice_sm_bd = sm_bd
                
                else:
                    slice_food_db = -404
                    slice_sm_bd = -404
        else:
            if fo_select == 0:
                fo_select = 1
    else:
        try:
            slice_food_db = slice_food_db
        except:
            slice_food_db = -404
        
        try:
            slice_sm_bd = slice_sm_bd
        except:
            slice_sm_bd = -404
        
        try:
            orderId = orderId
        except:
            orderId = None

        return slice_food_db, slice_sm_bd, orderId                               
                                        
def rtn_edit_finance(google_auth, fernet_key, rtn_database_url, order_concat, rtn_constants_dict, other_controls):
    payment_method = str(rtn_constants_dict["payment_selections"]).split(",")
    payment_include_others = eval(str(rtn_constants_dict["payment_selections_include_others"]).strip().capitalize())

    payment_columns = str(rtn_constants_dict["payment_columns"]).split(",")

    if payment_include_others:
        payment_method += ["其他"]
    else:
        pass

    fin_select = 0
    while fin_select != 1:
        try:
            payment_update = payment_update
        except:
            payment_update = -404
        
        if isinstance(payment_update, pd.DataFrame):
            option = ["保存并退出财务操作"]
        else:
            option = ["查看订单", "退出财务操作"]

        first_options = option_num(option)
        time.sleep(0.25)
        fin_select = option_limit(first_options, input("在这里输入>>>: "))

        if len(option) == 2:
            if fin_select == 0:
                orderId = rtn_select_order(order_concat=order_concat, rtn_constants_dict=rtn_constants_dict)

                if isinstance(orderId, str):
                    rtn_db, food_db, sm_bd, payment = rtn_order_filter(google_auth=google_auth, fernet_key=fernet_key, rtn_database_url=rtn_database_url, rtn_constants_dict=rtn_constants_dict, other_controls=other_controls, orderId=orderId)

                    try:
                        payment_update = payment_update
                    except:
                        payment_update = payment

                    if isinstance(payment_update, int):
                        if payment_update == -404:
                            payment_update  = payment

                    display_food_db, food_db, sm_bd_df, sm_bd, payment_info = rtn_food_order_parser(food_db=food_db, sm_bd=sm_bd, payment=payment_update, other_controls=other_controls)

                    second_select = 0
                    while second_select != 2:
                        print()
                        print()
                        print()
                        print("订单详情: ")
                        for index in range(len(rtn_db)):
                            for column in range(len(rtn_db.columns)):
                                print("{}: {}".format(rtn_db.columns[column], rtn_db.iloc[index, column]))
                        
                        print()
                        print("系统备注: 「预订除夕?」和「本地号码?」列, 1表是, 0表不是。")
                        print()
                        print()
                        if isinstance(display_food_db, pd.DataFrame):
                            display_food_db["点餐ID"] = display_food_db["点餐ID"].astype(str)
                            display_food_db["菜品ID"] = display_food_db["菜品ID"].astype(str)
                            print()
                            print("目前点餐详情: ")
                            prtdf(display_food_db.drop(["点餐ID", "订单ID", "菜品ID"], axis=1))
                            print("系统备注: 「套餐?」和「换菜?」列, 1表是, 0表不是。")
                        else:
                            print()
                            print("目前未点餐。")
                        
                        if isinstance(sm_bd_df, pd.DataFrame):
                            print()
                            print("套餐内菜肴详情: ")
                            prtdf(sm_bd_df.drop(["点餐ID"], axis=1))
                            print()
                        else:
                            print()
                        
                        print("财务详情: ")
                        print()
                        if payment_update.empty:
                            print("目前没有付款信息。")
                        else:
                            prtdf(payment_update.drop(["财务ID", "订单ID"], axis=1))
                        print()

                        for key, value in payment_info.items():
                            print("{} : {}".format(key, format(float(value), ".2f")))
                            
                        print()
                        print()

                        second_options = option_num(["添加付款记录", "删除付款记录", "返回上一菜单"])
                        time.sleep(0.25)
                        second_select = option_limit(second_options, input("在这里输入>>>: "))

                        if second_select == 0:
                            end_record_loop = False
                            while not end_record_loop:
                                
                                confirm_payment_mode = False
                                while not confirm_payment_mode:
                                    payment_methods = option_num(payment_method)
                                    time.sleep(0.25)
                                    payment_select = option_limit(payment_methods, input("在这里输入>>>: "))

                                    if payment_select != len(payment_method)-1:
                                        payment_mode = payment_method[payment_select]
                                    else:
                                        payment_mode = rtn_input_validation(rule="searchString", title="付款方式", space_removal=False, rtn_constants_dict=rtn_constants_dict)

                                    confirm_payment_mode = rtn_confirm_input(payment_mode)
                                
                                confirm_pay_value = False
                                while not confirm_pay_value:
                                    pay_value = rtn_input_validation(rule="foodPrice", title="付款金额", rtn_constants_dict=rtn_constants_dict)
                                    pay_value = float(format(float(pay_value), ".2f"))
                                    confirm_pay_value = rtn_confirm_input(pay_value)

                                pay_time = pd.to_datetime(dt.datetime.now())
                                paymentId = str(uuid.uuid1().hex)

                                sequence = [
                                    paymentId,
                                    orderId,
                                    pay_time,
                                    payment_mode,
                                    float(format(float(pay_value), ".2f"))]
                                
                                new_pay = {}
                                for index in range(len(payment_columns)):
                                    new_pay.update({ payment_columns[index] : [ sequence[index] ] })
                                
                                new_pay = pd.DataFrame(new_pay)
                                payment_concat = pd.concat([payment_update, new_pay], ignore_index=True)
                                payment_concat["付款时间"] = pd.to_datetime(payment_concat["付款时间"])
                                payment_concat.sort_values(by="付款时间", ascending=False, inplace=True, ignore_index=True)


                                display_food_db, food_db, sm_bd_df, sm_bd, uncof_payment_info = rtn_food_order_parser(food_db=food_db, sm_bd=sm_bd, payment=payment_concat, other_controls=other_controls)

                                print("更新后的财务详情预览: ")
                                print()
                                prtdf(payment_concat.drop(["财务ID", "订单ID"], axis=1))
                                print()
                                for key, value in uncof_payment_info.items():
                                    print("{} : {} ".format(key, format(float(value), ".2f")))
                                
                                print()
                                print()
                                confirm_save, end_record_loop = rtn_confirm_save("此记录" )

                            else:
                                if confirm_save:
                                    payment_update = payment_concat
                                    payment_info = uncof_payment_info
                                    print("财务记录已保存。")
                                else:
                                    print("财务记录未保存。")

                        elif second_select == 1:
                            if payment_update.empty:
                                print("无可删除的财务记录。")
                            else:
                                end_record_loop = False
                                while not end_record_loop:
                                    del_option = []
                                    paymentIdList = []
                                    for index in range(len(payment_update)):
                                        del_option += ["删除'金额: {}, 方式: {}'".format(payment_update.iloc[index, 4], payment_update.iloc[index, 3])]
                                        paymentIdList += ["{}".format(payment_update.iloc[index, 0])]
                                    
                                    del_options = option_num(del_option)
                                    time.sleep(0.25)
                                    del_select = option_limit(del_options, input("在这里输入>>>: "))

                                    payment_update["财务ID"] = payment_update["财务ID"].astype(str)

                                    print("你确定删除吗? ")
                                    confirm_del, end_record_loop = rtn_confirm_save("此删除")

                                else:
                                    if confirm_del:
                                        payment_update = payment_update.copy()
                                        payment_update = payment_update[payment_update["财务ID"] != paymentIdList[del_select]]
                                        payment_update["付款时间"] = pd.to_datetime(payment_update["付款时间"])
                                        payment_update.sort_values(by="付款时间", ascending=False, ignore_index=True, inplace=True)

                                        a, b, c, d, payment_info = rtn_food_order_parser(food_db=food_db, sm_bd=sm_bd, payment=payment_update, other_controls=other_controls)

                                        print("删除成功。")
                                    else:
                                        print("未删除。")
                    else:
                        try:
                            payment_update = payment_update
                        except:
                            payment_update = -404
                else:
                    payment_update = -404
                    payment_info_update = -404
        
        else:
            if fin_select == 0:
                fin_select = 1
    else:
        try:
            payment_update = payment_update
        except:
            payment_update = -404
        
        try:
            orderId = orderId
        except:
            orderId = None

        return payment_update, orderId

def rtn_edit_order(rtn_constants_dict, other_controls, google_auth, fernet_key, rtn_database_url):
    print()

    orderAttributeList = str(rtn_constants_dict["order_attribute"]).strip().split(",")
    comingCnyEve = pd.to_datetime(str(rtn_constants_dict["coming_cny_eve_date"]))
    tc = other_controls["tc"]
    tc["最小载客量"] = tc["最小载客量"].astype(int)
    tc["最大载客量"] = tc["最大载客量"].astype(int)

    order_columns = str(rtn_constants_dict["order_columns"]).split(",")

    comingCnyEve = pd.to_datetime(str(rtn_constants_dict["coming_cny_eve_date"]))

    rtn_db = rtn_fetch_database(google_auth=google_auth, fernet_key=fernet_key, rtn_database_url=rtn_database_url, rtn_constants_dict=rtn_constants_dict)["rtn_db"]
    rtn_db["预订除夕?"] = rtn_db["预订除夕?"].astype(int)
    rtn_db["订单属性"] = rtn_db["订单属性"].astype(str)
    rtn_db["订单号"] = rtn_db["订单号"].astype(str)

    orderNumber_allow_dup = eval(str(rtn_constants_dict["order_number_allow_dup"]).strip().capitalize())

    eo_select = 0
    while eo_select != 1:
        try:
            slice_rtn_db = slice_rtn_db
        except:
            slice_rtn_db = -404
        
        try:
            delete_orderId = delete_orderId
        except:
            delete_orderId = -404

        if isinstance(slice_rtn_db, pd.DataFrame):
            option = ["保存并退出编辑订单操作"]
        else:
            option = ["查看订单", "退出编辑订单操作"]
        
        first_options = option_num(option)
        time.sleep(0.25)
        eo_select = option_limit(first_options, input("在这里输入>>>: "))

        if len(option) == 2:
            if eo_select == 0:
                orderId = rtn_select_order(order_concat=rtn_db, rtn_constants_dict=rtn_constants_dict)

                if isinstance(orderId, str):
                    slice_rtn_db, food_db, sm_bd, payment = rtn_order_filter(google_auth=google_auth, fernet_key=fernet_key, rtn_database_url=rtn_database_url, rtn_constants_dict=rtn_constants_dict, other_controls=other_controls, orderId=orderId)

                    second_select = 0
                    while second_select != 10:
                        print()
                        print()
                        print("订单详情: ")
                        for index in range(len(slice_rtn_db)):
                            for column in range(len(slice_rtn_db.columns)):
                                print("{}: {}".format(slice_rtn_db.columns[column], slice_rtn_db.iloc[index, column]))

                        print()
                        print("系统备注: 「预订除夕?」和「本地号码?」列, 1表是, 0表不是。")
                        print()
                        slice_rtn_db = rtn_datetime_column_convert(slice_rtn_db)
                        slice_rtn_db["订单ID"] = slice_rtn_db["订单ID"].astype(str)
                        slice_rtn_db["桌台"] = slice_rtn_db["桌台"].astype(str)


                        second_options = option_num(["修改姓名", "修改订单号", "修改电话", "修改预订时间/轮数", "修改人数/载客量", "修改桌台", "修改订单备注", "修改订单状态", "修改订单属性", "删除订单", "完成变更并返回上一菜单"])
                        time.sleep(0.25)
                        second_select = option_limit(second_options, input("在这里输入>>>: "))

                        if second_select == 0:
                            try:
                                delete_orderId = delete_orderId
                            except:
                                delete_orderId  = -404

                            if isinstance(delete_orderId, int):
                                current_customer_name = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["姓名"].values[0])
                                print("目前客人名字: {}".format(current_customer_name))

                                end_record_loop = False
                                while not end_record_loop:
                                    customerName = rtn_input_validation(rule="searchString", title="客户名", space_removal=False, rtn_constants_dict=rtn_constants_dict)

                                    print()
                                    print()
                                    print("你确定把客人的名字从'{}'改成'{}'吗? ".format(current_customer_name, customerName))
                                    confirm_save, end_record_loop = rtn_confirm_save(customerName)
                                else:
                                    if confirm_save:
                                        outletCode = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["门店代号"].values[0])
                                        orderCreatedTime = pd.to_datetime(str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单创建时间"].values[0]))
                                        orderAttribute = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单属性"].values[0])
                                        reserveTime = pd.to_datetime(str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["预订时间"].values[0]))
                                        isCnyEve = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["预订除夕?"].values[0])
                                        orderNumber = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单号"].values[0])
                                        phone = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["电话"].values[0])
                                        isLocalPhone = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["本地号码?"].values[0])
                                        round = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["轮数"].values[0])
                                        adultPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["成人人数"].values[0])
                                        childPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["儿童人数"].values[0])
                                        toddlerPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["幼儿人数"].values[0])
                                        totalPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["载客量"].values[0])
                                        tableAssign = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["桌台"].values[0])
                                        remark = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["备注"].values[0])
                                        orderStatus = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单状态"].values[0])

                                        sequence = [
                                            outletCode,
                                            orderId,
                                            orderCreatedTime,
                                            orderAttribute,
                                            reserveTime,
                                            isCnyEve,
                                            customerName,
                                            orderNumber,
                                            phone,
                                            isLocalPhone,
                                            round,
                                            adultPax,
                                            childPax,
                                            toddlerPax,
                                            totalPax,
                                            tableAssign,
                                            remark,
                                            orderStatus]

                                        modify = {}
                                        for index in range(len(order_columns)):
                                            modify.update({ order_columns[index] : [sequence[index]] })
                                        
                                        modify = pd.DataFrame(modify)

                                        slice_rtn_db = modify
                                        print("修改成功。")
                                    else:
                                        print("未修改。")

                            else:
                                print()
                                print()
                                print("你已经选择删除这个订单了。")
                                time.sleep(0.25)
                                input("按回车键继续>>>: ")
                                print()
                                print()

                        elif second_select == 1:
                            try:
                                delete_orderId = delete_orderId
                            except:
                                delete_orderId  = -404

                            if isinstance(delete_orderId, int):
                                current_order_number = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单号"].values[0])
                                print("目前订单号: {}".format(current_order_number))

                                end_record_loop = False
                                while not end_record_loop:
                                    orderNumberOptions = option_num(["自动生成随机订单号", "手动输入订单号"])
                                    time.sleep(0.25)
                                    orderNumberSelect = option_limit(orderNumberOptions, input("在这里输入>>>: "))

                                    if orderNumberSelect == 0:
                                        orderNumber = str(uuid.uuid1().hex)
                                    elif orderNumberSelect == 1:
                                        print("注意! 订单号可以通过修改后而重复! ")
                                        orderNumber = rtn_input_validation(rule="orderNumber", title="订单号", space_removal=True, rtn_constants_dict=rtn_constants_dict)

                                    if not orderNumber_allow_dup:
                                        if str(orderNumber) in rtn_db["订单号"].values.tolist():
                                            print("订单号重复! 请重新输入订单号! ")
                                            end_record_loop = False
                                    
                                    print()
                                    print()
                                    print("你确定把订单号从'{}'改成'{}'吗? ".format(current_order_number, orderNumber))
                                    confirm_save, end_record_loop = rtn_confirm_save(orderNumber)
                                
                                else:
                                    if confirm_save:
                                        outletCode = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["门店代号"].values[0])
                                        orderCreatedTime = pd.to_datetime(str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单创建时间"].values[0]))
                                        orderAttribute = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单属性"].values[0])
                                        reserveTime = pd.to_datetime(str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["预订时间"].values[0]))
                                        isCnyEve = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["预订除夕?"].values[0])
                                        customerName = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["姓名"].values[0])
                                        phone = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["电话"].values[0])
                                        isLocalPhone = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["本地号码?"].values[0])
                                        round = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["轮数"].values[0])
                                        adultPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["成人人数"].values[0])
                                        childPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["儿童人数"].values[0])
                                        toddlerPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["幼儿人数"].values[0])
                                        totalPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["载客量"].values[0])
                                        tableAssign = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["桌台"].values[0])
                                        remark = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["备注"].values[0])
                                        orderStatus = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单状态"].values[0])

                                        sequence = [
                                            outletCode,
                                            orderId,
                                            orderCreatedTime,
                                            orderAttribute,
                                            reserveTime,
                                            isCnyEve,
                                            customerName,
                                            orderNumber,
                                            phone,
                                            isLocalPhone,
                                            round,
                                            adultPax,
                                            childPax,
                                            toddlerPax,
                                            totalPax,
                                            tableAssign,
                                            remark,
                                            orderStatus]

                                        modify = {}
                                        for index in range(len(order_columns)):
                                            modify.update({ order_columns[index] : [sequence[index]] })
                                        
                                        modify = pd.DataFrame(modify)

                                        slice_rtn_db = modify
                                        print("修改成功。")
                                    else:
                                        print("未修改。")

                            else:
                                print()
                                print()
                                print("你已经选择删除这个订单了。")
                                time.sleep(0.25)
                                input("按回车键继续>>>: ")
                                print()
                                print()

                        elif second_select == 2:
                            try:
                                delete_orderId = delete_orderId
                            except:
                                delete_orderId  = -404

                            if isinstance(delete_orderId, int):
                                currPhone = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["电话"].values[0])
                                print("目前电话号码: {}".format(currPhone))

                                end_record_loop = False
                                while not end_record_loop:
                                    phoneOptionList = ["本地号码", "外国号码", "无预留号码"]
                                    phoneOptions = option_num(phoneOptionList)
                                    time.sleep(0.25)
                                    phoneOptionSelect = option_limit(phoneOptions, input("在这里输入>>>: "))

                                    if phoneOptionSelect == 0:
                                        customerPhone = rtn_input_validation(rule="localPhone", title="本地号码")
                                    
                                    elif phoneOptionSelect == 1:
                                        print("外国号码请添加国际代码,但无需'+'号。")
                                        customerPhone = rtn_input_validation(rule="searchString", title="外国号码", remove_spaces=True, rtn_constants_dict=rtn_constants_dict)

                                        if "+" in customerPhone:
                                            customerPhone = customerPhone.replace("+", "")
                                            print("'+'已自动移除, 无需添加'+'号。")
                                        
                                    
                                    else:
                                        print("无预留电话将以'-'呈现。")
                                        customerPhone = "-"

                                    if phoneOptionList[phoneOptionSelect] == phoneOptionList[0]:
                                        isLocalPhone = 1
                                    else:
                                        isLocalPhone = 0
                                    
                                    print()
                                    print()
                                    print("你确定把电话号码从'{}'换成'{}'吗? ".format(currPhone, customerPhone))
                                    confirm_save, end_record_loop = rtn_confirm_save(customerPhone)
                                
                                else:
                                    if confirm_save:
                                        outletCode = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["门店代号"].values[0])
                                        orderCreatedTime = pd.to_datetime(str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单创建时间"].values[0]))
                                        orderAttribute = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单属性"].values[0])
                                        reserveTime = pd.to_datetime(str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["预订时间"].values[0]))
                                        isCnyEve = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["预订除夕?"].values[0])
                                        customerName = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["姓名"].values[0])
                                        orderNumber = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单号"].values[0])
                                        phone = customerPhone
                                        round = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["轮数"].values[0])
                                        adultPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["成人人数"].values[0])
                                        childPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["儿童人数"].values[0])
                                        toddlerPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["幼儿人数"].values[0])
                                        totalPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["载客量"].values[0])
                                        tableAssign = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["桌台"].values[0])
                                        remark = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["备注"].values[0])
                                        orderStatus = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单状态"].values[0])

                                        sequence = [
                                            outletCode,
                                            orderId,
                                            orderCreatedTime,
                                            orderAttribute,
                                            reserveTime,
                                            isCnyEve,
                                            customerName,
                                            orderNumber,
                                            phone,
                                            isLocalPhone,
                                            round,
                                            adultPax,
                                            childPax,
                                            toddlerPax,
                                            totalPax,
                                            tableAssign,
                                            remark,
                                            orderStatus]

                                        modify = {}
                                        for index in range(len(order_columns)):
                                            modify.update({ order_columns[index] : [sequence[index]] })
                                        
                                        modify = pd.DataFrame(modify)

                                        slice_rtn_db = modify
                                        print("修改成功。")
                                    else:
                                        print("未修改。")

                            else:
                                print()
                                print()
                                print("你已经选择删除这个订单了。")
                                time.sleep(0.25)
                                input("按回车键继续>>>: ")
                                print()
                                print()   
                        
                        elif second_select == 3:
                            try:
                                delete_orderId = delete_orderId
                            except:
                                delete_orderId  = -404

                            if isinstance(delete_orderId, int):
                                CnyEveDate = dt.datetime(year=comingCnyEve.year, month=comingCnyEve.month, day=comingCnyEve.day)

                                currReserveDatetime = pd.to_datetime(str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["预订时间"].values[0]))
                                print("目前预订时间: {}".format(currReserveDatetime.strftime("%Y-%m-%d %H:%M")))

                                end_record_loop = False
                                while not end_record_loop:
                                    reserveTime = rtn_edit_datetime()
                                
                                    reserveDate = dt.datetime(year=reserveTime.year, month=reserveTime.month, day=reserveTime.day)

                                    if CnyEveDate == reserveDate:
                                        isCnyEve = 1
                                    else:
                                        isCnyEve = 0

                                    round_time_param = str(rtn_constants_dict["cny_eve_round_time_param"]).split(",")
                
                                    round_time_dt_bind = {}
                                    for t in range(len(round_time_param)):
                                        dt_format = dt.datetime.strptime(round_time_param[t], "%Y-%m-%d %H:%M")
                                        round_sequence = "第{}轮".format(t+1)
                                        round_time_dt_bind.update({ dt_format : round_sequence  })
                                    
                                    if reserveTime in list(round_time_dt_bind.keys()):
                                        round = round_time_dt_bind[reserveTime]
                                    else:
                                        round = "无效"
                                    
                                    print()
                                    print()
                                    print("影响后的对象: ")
                                    print("预订时间: {}".format(reserveTime.strftime("%Y-%m-%d %H:%M")))
                                    print("预订除夕?(0表不是, 1表是): {}".format(isCnyEve))
                                    print("轮数: {}".format(round))
                                    print()
                                    print()
                                    print("你确定把预订时间从'{}'改成'{}'吗? ".format(currReserveDatetime.strftime("%Y-%m-%d %H:%M"), reserveTime.strftime("%Y-%m-%d %H:%M")))
                                    confirm_save, end_record_loop = rtn_confirm_save(reserveTime.strftime("%Y-%m-%d %H:%M"))
                                else:
                                    if confirm_save:
                                        orderAttribute = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单属性"].values[0]) 
                                        print("修改预订时间/轮数后, 以下对象将重置: ")
                                        if orderAttribute == "堂食":
                                            tableAssign = "未排位"
                                        else:
                                            tableAssign = "无效"
                                        
                                        print("桌台重置为'{}'。".format(tableAssign))

                                        outletCode = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["门店代号"].values[0])
                                        orderCreatedTime = pd.to_datetime(str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单创建时间"].values[0]))                
                                        customerName = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["姓名"].values[0])
                                        orderNumber = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单号"].values[0])
                                        phone = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["电话"].values[0])
                                        adultPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["成人人数"].values[0])
                                        childPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["儿童人数"].values[0])
                                        toddlerPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["幼儿人数"].values[0])
                                        totalPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["载客量"].values[0])
                                        remark = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["备注"].values[0])
                                        orderStatus = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单状态"].values[0])
                                        isLocalPhone = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["本地号码?"].values[0])

                                        sequence = [
                                            outletCode,
                                            orderId,
                                            orderCreatedTime,
                                            orderAttribute,
                                            reserveTime,
                                            isCnyEve,
                                            customerName,
                                            orderNumber,
                                            phone,
                                            isLocalPhone,
                                            round,
                                            adultPax,
                                            childPax,
                                            toddlerPax,
                                            totalPax,
                                            tableAssign,
                                            remark,
                                            orderStatus]

                                        modify = {}
                                        for index in range(len(order_columns)):
                                            modify.update({ order_columns[index] : [sequence[index]] })
                                        
                                        modify = pd.DataFrame(modify)

                                        slice_rtn_db = modify
                                        print("修改成功。")
                                    else:
                                        print("未修改。")

                            else:
                                print()
                                print()
                                print("你已经选择删除这个订单了。")
                                time.sleep(0.25)
                                input("按回车键继续>>>: ")
                                print()
                                print()

                        elif second_select == 4:
                            try:
                                delete_orderId = delete_orderId
                            except:
                                delete_orderId = -404
                            
                            if isinstance(delete_orderId, int):
                                orderAttribute = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单属性"].values[0])

                                if orderAttribute == "堂食":
                                    currAdultPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["成人人数"].values[0])
                                    currChildPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["儿童人数"].values[0])
                                    currToddlerPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["幼儿人数"].values[0])
                                    currTotalPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["载客量"].values[0])

                                    print("目前人数详情: ")
                                    print("成人人数: {}".format(currAdultPax))
                                    print("儿童人数: {}".format(currChildPax))
                                    print("幼儿人数: {}".format(currToddlerPax))
                                    print("载客量: {}".format(currTotalPax))
                                    print()

                                    end_record_loop = False
                                    while not end_record_loop:
                                        adultPax = rtn_input_validation(rule="pax", title="成人人数", other_controls=other_controls)
                                        childPax = rtn_input_validation(rule="notAdultPax", title="儿童人数", other_controls=other_controls)
                                        toddlerPax = rtn_input_validation(rule="notAdultPax", title="幼儿人数", other_controls=other_controls)

                                        confirm_totalPax = False
                                        while not confirm_totalPax:
                                            totalPax = rtn_input_validation(rule="pax", title="载客量", other_controls=other_controls)

                                            if totalPax < adultPax:
                                                print("载客量不能小于成人人数, 请重新输入! ")
                                                confirm_totalPax = False
                                            else:
                                                confirm_totalPax = True
                                        
                                        print()
                                        print()
                                        print("更改后详情: ")
                                        print("成人人数: {}".format(adultPax))
                                        print("儿童人数: {}".format(childPax))
                                        print("幼儿人数: {}".format(toddlerPax))
                                        print("载客量: {}".format(totalPax))
                                        print()
                                        print()
                                        print("你确定这些更改吗? ")
                                        confirm_save, end_record_loop = rtn_confirm_save("这些更改")
                                    else:
                                        if confirm_save:
                                            outletCode = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["门店代号"].values[0])
                                            orderCreatedTime = pd.to_datetime(str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单创建时间"].values[0]))
                                            orderAttribute = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单属性"].values[0]) 
                                            reserveTime = pd.to_datetime(str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["预订时间"].values[0]))
                                            isCnyEve = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["预订除夕?"].values[0])                               
                                            customerName = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["姓名"].values[0])
                                            orderNumber = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单号"].values[0])
                                            phone = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["电话"].values[0])
                                            isLocalPhone = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["本地号码?"].values[0])
                                            round = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["轮数"].values[0])
                                            tableAssign = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["桌台"].values[0])
                                            remark = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["备注"].values[0])
                                            orderStatus = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单状态"].values[0])

                                            sequence = [
                                                outletCode,
                                                orderId,
                                                orderCreatedTime,
                                                orderAttribute,
                                                reserveTime,
                                                isCnyEve,
                                                customerName,
                                                orderNumber,
                                                phone,
                                                isLocalPhone,
                                                round,
                                                adultPax,
                                                childPax,
                                                toddlerPax,
                                                totalPax,
                                                tableAssign,
                                                remark,
                                                orderStatus]

                                            modify = {}
                                            for index in range(len(order_columns)):
                                                modify.update({ order_columns[index] : [sequence[index]] })
                                            
                                            modify = pd.DataFrame(modify)

                                            slice_rtn_db = modify
                                            print("修改成功。")
                                        else:
                                            print("未修改。")

                                else:
                                    print("打包订单不能修改人数。")
                                    adultPax = 0
                                    childPax = 0
                                    toddlerPax = 0
                                    totalPax = 0

                                    outletCode = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["门店代号"].values[0])
                                    orderCreatedTime = pd.to_datetime(str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单创建时间"].values[0]))
                                    orderAttribute = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单属性"].values[0])
                                    isCnyEve = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["预订除夕?"].values[0])
                                    reserveTime = pd.to_datetime(str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["预订时间"].values[0]))                               
                                    customerName = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["姓名"].values[0])
                                    orderNumber = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单号"].values[0])
                                    phone = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["电话"].values[0])
                                    isLocalPhone = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["本地号码?"].values[0])
                                    round = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["轮数"].values[0])
                                    tableAssign = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["桌台"].values[0])
                                    remark = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["备注"].values[0])
                                    orderStatus = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单状态"].values[0])

                                    sequence = [
                                            outletCode,
                                            orderId,
                                            orderCreatedTime,
                                            orderAttribute,
                                            reserveTime,
                                            isCnyEve,
                                            customerName,
                                            orderNumber,
                                            phone,
                                            isLocalPhone,
                                            round,
                                            adultPax,
                                            childPax,
                                            toddlerPax,
                                            totalPax,
                                            tableAssign,
                                            remark,
                                            orderStatus]

                                    modify = {}
                                    for index in range(len(order_columns)):
                                        modify.update({ order_columns[index] : [sequence[index]] })
                                    
                                    modify = pd.DataFrame(modify)

                                    slice_rtn_db = modify

                            else:
                                print()
                                print()
                                print("你已经选择删除这个订单了。")
                                time.sleep(0.25)
                                input("按回车键继续>>>: ")
                                print()
                                print()

                        elif second_select == 5:
                            try:
                                delete_orderId = delete_orderId
                            except:
                                delete_orderId = -404
                            
                            if isinstance(delete_orderId, int):
                                orderAttribute = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单属性"].values[0])
                                reserveTime = pd.to_datetime(str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["预订时间"].values[0]))
                                round = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["轮数"].values[0])
                                totalPax = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["载客量"].values[0])

                                end_record_loop = False
                                while not end_record_loop:
                                    if orderAttribute != orderAttributeList[0]:
                                        print("打包订单无需桌台")
                                        tableAssign = "无效"
                                        confirm_save = True
                                        end_record_loop = True

                                    else:
                                        currTable = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["桌台"].values[0])
                                        print("目前桌台: {}".format(currTable))

                                        confirm_table_assign = False
                                        while not confirm_table_assign:
                                            tableAssignOptions = option_num(["暂不排位", "排位"])
                                            time.sleep(0.25)
                                            tableAssignSelect = option_limit(tableAssignOptions, input("在这里输入>>>: "))

                                            if tableAssignSelect == 0:
                                                tableAssign = "未排位"
                                                confirm_table_assign = True
                                            
                                            elif tableAssignSelect == 1:
                                                table_occupancy, unassigned_df = rtn_table_occupancy(rtn_db=rtn_db, round=round, reserveTime=reserveTime, other_controls=other_controls)
                                                table_occupancy["桌名"] = table_occupancy["桌名"].astype(str)

                                                print("目前占位状况: ")
                                                prtdf(table_occupancy)
                                                print()
                                                print("现在开始选择你要排的台位...")
                                                print()
                                                tableId = rtn_select_table(rtn_constants_dict=rtn_constants_dict, other_controls=other_controls)

                                                if not isinstance(tableId, str):
                                                    print("选台位失败, 请重新选择! ")
                                                    confirm_table_assign = False
                                                else:
                                                    tableName = str(tc[tc["桌子ID"] == tableId]["桌名"].values[0])

                                                    if int(totalPax) <= int(tc[tc["桌子ID"] == tableId]["最大载客量"].values[0]):
                                                        if int(totalPax) >= int(tc[tc["桌子ID"] == tableId]["最小载客量"].values[0]):
                                                            if str(tableName) in slice_rtn_db["桌台"].values:
                                                                print("排位失败, 该桌台在当天当轮已被占用! ")
                                                                print("请重新选择! ")
                                                                confirm_table_assign = False
                                                            else:
                                                                confirm_table_assign = True
                                                        else:
                                                            print("排位失败, 该桌台太大了! ")
                                                            print("请重新选择! ")
                                                            confirm_table_assign = False
                                                    else:
                                                        print("排位失败, 该桌台太小了! ")
                                                        print("请重新选择! ")
                                                        confirm_table_assign = False
                                            
                                        tableAssign = tableName
                                        print()
                                        print()
                                        print("你确定把桌台'{}'改成'{}'吗? ".format(currTable, tableAssign))
                                        confirm_save, end_record_loop = rtn_confirm_save("此改动")
                                else:
                                    if confirm_save:
                                        outletCode = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["门店代号"].values[0])
                                        orderCreatedTime = pd.to_datetime(str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单创建时间"].values[0]))
                                        orderAttribute = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单属性"].values[0])
                                        reserveTime = pd.to_datetime(str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["预订时间"].values[0]))
                                        isCnyEve = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["预订除夕?"].values[0])
                                        customerName = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["姓名"].values[0])
                                        orderNumber = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单号"].values[0])
                                        phone = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["电话"].values[0])
                                        isLocalPhone = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["本地号码?"].values[0])
                                        round = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["轮数"].values[0])
                                        adultPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["成人人数"].values[0])
                                        childPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["儿童人数"].values[0])
                                        toddlerPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["幼儿人数"].values[0])
                                        totalPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["载客量"].values[0])
                                        remark = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["备注"].values[0])
                                        orderStatus = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单状态"].values[0])

                                        sequence = [
                                            outletCode,
                                            orderId,
                                            orderCreatedTime,
                                            orderAttribute,
                                            reserveTime,
                                            isCnyEve,
                                            customerName,
                                            orderNumber,
                                            phone,
                                            isLocalPhone,
                                            round,
                                            adultPax,
                                            childPax,
                                            toddlerPax,
                                            totalPax,
                                            tableAssign,
                                            remark,
                                            orderStatus]

                                        modify = {}
                                        for index in range(len(order_columns)):
                                            modify.update({ order_columns[index] : [sequence[index]] })
                                        
                                        modify = pd.DataFrame(modify)

                                        slice_rtn_db = modify
                                        print("桌台更改完成。")
                                    else:
                                        print("桌台更改完成。")
                                        print("未修改。")

                            else:
                                print()
                                print()
                                print("你已经选择删除这个订单了。")
                                time.sleep(0.25)
                                input("按回车键继续>>>: ")
                                print()
                                print()

                        elif second_select == 6:
                            try:
                                delete_orderId = delete_orderId
                            except:
                                delete_orderId  = -404

                            if isinstance(delete_orderId, int):
                                currRemark = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["备注"].values[0])
                                print("目前备注: {}".format(currRemark))

                                end_record_loop = False
                                while not end_record_loop:
                                    remarkOption = option_num(["输入备注", "无备注"])
                                    time.sleep(0.25)
                                    remarkOptionSelect = option_limit(remarkOption, input("在这里输入>>>: "))

                                    if remarkOptionSelect == 0:
                                        remark = rtn_input_validation(rule="remark", title="备注", space_removal=False, rtn_constants_dict=rtn_constants_dict)

                                        if len(remark) == 0:
                                            print("无备注是吧? 自动用'-'呈现。")
                                    else:
                                        print("无备注将用'-'呈现。")
                                        remark = "-"
                                    
                                    print("你确定把替换为这些备注内容吗? ")
                                    print("原来的备注: ")
                                    print(currRemark)
                                    print()
                                    print()
                                    print("更改后的备注: ")
                                    print(remark)
                                    print()
                                    print()
                                    print()
                                    confirm_save, end_record_loop = rtn_confirm_save("替换为这些备注内容")
                                else:
                                    if confirm_save:
                                        outletCode = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["门店代号"].values[0])
                                        orderCreatedTime = pd.to_datetime(str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单创建时间"].values[0]))
                                        orderAttribute = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单属性"].values[0])
                                        reserveTime = pd.to_datetime(str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["预订时间"].values[0]))
                                        isCnyEve = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["预订除夕?"].values[0])
                                        customerName = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["姓名"].values[0])
                                        orderNumber = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单号"].values[0])
                                        phone = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["电话"].values[0])
                                        isLocalPhone = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["本地号码?"].values[0])
                                        round = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["轮数"].values[0])
                                        adultPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["成人人数"].values[0])
                                        childPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["儿童人数"].values[0])
                                        toddlerPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["幼儿人数"].values[0])
                                        totalPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["载客量"].values[0])
                                        tableAssign = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["桌台"].values[0])
                                        orderStatus = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单状态"].values[0])

                                        sequence = [
                                            outletCode,
                                            orderId,
                                            orderCreatedTime,
                                            orderAttribute,
                                            reserveTime,
                                            isCnyEve,
                                            customerName,
                                            orderNumber,
                                            phone,
                                            isLocalPhone,
                                            round,
                                            adultPax,
                                            childPax,
                                            toddlerPax,
                                            totalPax,
                                            tableAssign,
                                            remark,
                                            orderStatus]

                                        modify = {}
                                        for index in range(len(order_columns)):
                                            modify.update({ order_columns[index] : [sequence[index]] })
                                        
                                        modify = pd.DataFrame(modify)

                                        slice_rtn_db = modify
                                        print("修改成功。")
                                    else:
                                        print("未修改。")

                            else:
                                print()
                                print()
                                print("你已经选择删除这个订单了。")
                                time.sleep(0.25)
                                input("按回车键继续>>>: ")
                                print()
                                print()

                        elif second_select == 7:
                            try:
                                delete_orderId = delete_orderId
                            except:
                                delete_orderId  = -404
                            
                            if isinstance(delete_orderId, int):
                                display_food_db, food_db, sm_bd_df, sm_bd, payment_info = rtn_food_order_parser(food_db=food_db, sm_bd=sm_bd, payment=payment, other_controls=other_controls)

                                currOrderStatus = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单状态"].values[0])
                                print("目前订单状态: {}".format(currOrderStatus))
                                print()

                                option = str(rtn_constants_dict["order_status"]).split(",")

                                end_record_loop = False
                                while not end_record_loop:
                                    if payment_info["需付金额"] == 0:
                                        option = option
                                    else:
                                        print("未付全款的订单无法完结。")
                                        option.remove("完结")
                                    
                                    options = option_num(option)
                                    time.sleep(0.25)
                                    option_select = option_limit(options, input("在这里输入>>>: "))
                                
                                    orderStatus = option[option_select]
                                    print()
                                    print()
                                    print("你确定把该订单的状态从'{}'改成'{}'吗? ".format(currOrderStatus, orderStatus))
                                    confirm_save, end_record_loop = rtn_confirm_save("'{}'".format(orderStatus))
                                
                                else:
                                    if confirm_save:
                                        outletCode = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["门店代号"].values[0])
                                        orderCreatedTime = pd.to_datetime(str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单创建时间"].values[0]))
                                        orderAttribute = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单属性"].values[0])
                                        reserveTime = pd.to_datetime(str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["预订时间"].values[0]))
                                        isCnyEve = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["预订除夕?"].values[0])
                                        customerName = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["姓名"].values[0])
                                        orderNumber = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单号"].values[0])
                                        phone = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["电话"].values[0])
                                        isLocalPhone = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["本地号码?"].values[0])
                                        round = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["轮数"].values[0])
                                        adultPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["成人人数"].values[0])
                                        childPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["儿童人数"].values[0])
                                        toddlerPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["幼儿人数"].values[0])
                                        totalPax = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["载客量"].values[0])
                                        tableAssign = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["桌台"].values[0])
                                        remark = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["备注"].values[0])

                                        sequence = [
                                            outletCode,
                                            orderId,
                                            orderCreatedTime,
                                            orderAttribute,
                                            reserveTime,
                                            isCnyEve,
                                            customerName,
                                            orderNumber,
                                            phone,
                                            isLocalPhone,
                                            round,
                                            adultPax,
                                            childPax,
                                            toddlerPax,
                                            totalPax,
                                            tableAssign,
                                            remark,
                                            orderStatus]

                                        modify = {}
                                        for index in range(len(order_columns)):
                                            modify.update({ order_columns[index] : [sequence[index]] })
                                        
                                        modify = pd.DataFrame(modify)

                                        slice_rtn_db = modify
                                        print("修改成功。")
                                    else:
                                        print("未修改。")

                            else:
                                print()
                                print()
                                print("你已经选择删除这个订单了。")
                                time.sleep(0.25)
                                input("按回车键继续>>>: ")
                                print()
                                print()

                        elif second_select == 8:
                            try:
                                delete_orderId = delete_orderId
                            except:
                                delete_orderId  = -404
                            
                            if isinstance(delete_orderId, int):
                                display_food_db, food_db, sm_bd_df, sm_bd, payment_info = rtn_food_order_parser(food_db=food_db, sm_bd=sm_bd, payment=payment, other_controls=other_controls)

                                food_db["订单ID"] = food_db["订单ID"].astype(str)
                                food_db["菜品属性"] = food_db["菜品属性"].astype(str)
                                food_db = food_db[food_db["订单ID"] == orderId]
                                
                                orderAttribute = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单属性"].values[0])

                                if orderAttribute == "堂食":
                                    if "堂食" in food_db["菜品属性"].values:
                                        print("发现有菜品属性为'堂食', 无法更改订单属性。")
                                        print("如要更改订单属性, 请把该订单下所有的菜品属性改成'打包'后再回来更改。")
                                        print()
                                        canChangeAttri = False
                                    else:
                                        print("注意该订单下有打包的菜品。")
                                        canChangeAttri = True
                                else:
                                    canChangeAttri = True
                                
                                if canChangeAttri:
                                    end_record_loop = False
                                    while not end_record_loop:
                                        print()
                                        currAttri = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单属性"].values[0])
                                        print("目前订单属性: {}".format(currAttri))
                                        print()

                                        options = option_num(orderAttributeList)
                                        time.sleep(0.25)
                                        option_select = option_limit(options, input("在这里输入>>>: "))

                                        orderAttribute = orderAttributeList[option_select]
                                        print()
                                        print()

                                        print("你确定把订单属性从'{}'改成'{}'吗? ".format(currAttri, orderAttribute))
                                        confirm_save, end_record_loop = rtn_confirm_save(orderAttribute)
                                    else:
                                        if confirm_save:
                                            print("更改订单属性后, 以下对象将重置: ")
                                            print("成人人数重置为0。")
                                            print("儿童人数重置为0。")
                                            print("幼儿人数重置为0。")
                                            print("载客量重置为0。")

                                            if orderAttribute == "堂食":
                                                tableAssign = "未排位"
                                                print("桌台重置为'未排位'。")
                                            else:
                                                tableAssign = "无效"
                                                print("桌台重置为'无效'。")
                                            
                                            outletCode = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["门店代号"].values[0])
                                            orderCreatedTime = pd.to_datetime(str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单创建时间"].values[0]))
                                            reserveTime = pd.to_datetime(str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["预订时间"].values[0]))
                                            isCnyEve = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["预订除夕?"].values[0])
                                            customerName = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["姓名"].values[0])
                                            orderNumber = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单号"].values[0])
                                            phone = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["电话"].values[0])
                                            isLocalPhone = int(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["本地号码?"].values[0])
                                            round = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["轮数"].values[0])
                                            adultPax = 0
                                            childPax = 0
                                            toddlerPax = 0
                                            totalPax = 0
                                            remark = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["备注"].values[0])
                                            orderStatus = str(slice_rtn_db[slice_rtn_db["订单ID"] == orderId]["订单状态"].values[0])

                                            sequence = [
                                                outletCode,
                                                orderId,
                                                orderCreatedTime,
                                                orderAttribute,
                                                reserveTime,
                                                isCnyEve,
                                                customerName,
                                                orderNumber,
                                                phone,
                                                isLocalPhone,
                                                round,
                                                adultPax,
                                                childPax,
                                                toddlerPax,
                                                totalPax,
                                                tableAssign,
                                                remark,
                                                orderStatus]

                                            modify = {}
                                            for index in range(len(order_columns)):
                                                modify.update({ order_columns[index] : [sequence[index]] })
                                            
                                            modify = pd.DataFrame(modify)

                                            slice_rtn_db = modify
                                            print("修改成功。")
                                        else:
                                            print("未修改。")

                            else:
                                print()
                                print()
                                print("你已经选择删除这个订单了。")
                                time.sleep(0.25)
                                input("按回车键继续>>>: ")
                                print()
                                print()

                        elif second_select == 9:
                            print()
                            print()
                            print("警告: 如果你删除了订单, 下一次新建的订单号可能会取代此订单号。")
                            print("建议你把订单状态改成'取消', 而不是删除订单。")
                            print()
                            print("你确定删除这个订单吗? ")
                            option = ["是的, 删除这个订单", "不是, 别删除这个订单"]
                            options = option_num(option)
                            time.sleep(0.25)
                            select = option_limit(options, input("在这里输入>>>: "))

                            if select == 0:
                                delete_orderId = str(orderId)
                                print("好的, 即将删除。")
                                print("如果想改变主意, 重新回来这里改变选择。")
                            else:
                                delete_orderId = -404
                                print("好的, 不会删除。")
                    else:
                        try:
                            slice_rtn_db = slice_rtn_db
                        except:
                            slice_rtn_db = -404
                        
                        try:
                            delete_orderId = delete_orderId
                        except:
                            delete_orderId = -404
                
            else:
                try:
                    slice_rtn_db = slice_rtn_db
                except:
                    slice_rtn_db = -404
                
                try:
                    delete_orderId = delete_orderId
                except:
                    delete_orderId = -404
        else:
            if eo_select == 0:
                eo_select = 1
    else:
        try:
            slice_rtn_db = slice_rtn_db
        except:
            slice_rtn_db = -404
        
        try:
            delete_orderId = delete_orderId
        except:
            delete_orderId = -404

        return slice_rtn_db, delete_orderId        

def rtn_table_occupancy(rtn_db, round, reserveTime, other_controls):
    tc = other_controls["tc"]
    tc["桌名"] = tc["桌名"].astype(str)
    tc["最小载客量"] = tc["最小载客量"].astype(int)
    tc["最大载客量"] = tc["最大载客量"].astype(int)

    rtn_db["订单属性"] = rtn_db["订单属性"].astype(str)
    rtn_db["轮数"] = rtn_db["轮数"].astype(str)
    rtn_db["预订时间"] = pd.to_datetime(rtn_db["预订时间"])
    rtn_db["桌台"] = rtn_db["桌台"].astype(str)
    rtn_db["载客量"] = rtn_db["载客量"].astype(int)
    rtn_db["订单ID"] = rtn_db["订单ID"].astype(str)

    rtn_db = rtn_db[rtn_db["订单属性"] == "堂食"]
    filter = (rtn_db["轮数"] == round) & (rtn_db["预订时间"] == reserveTime)
    rtn_db = rtn_db[filter]

    unassigned_orderId = rtn_db[rtn_db["桌台"] == "未排位"]["订单ID"].values.astype(str).tolist()

    tableNameArr = []
    tableOccupancy = []
    tableOccupyByOrderNumber = []
    OrderTotalPax = []
    for index in range(len(tc)):
        tableName = str(tc.iloc[index, 1])
        tableMin = int(tc.iloc[index, 2])
        tableMax = int(tc.iloc[index, 3])

        filter = (rtn_db["桌台"] == tableName)

        tableNameArr += [tableName]
        tableOccupancy += ["{}-{}人".format(tableMin, tableMax)]
        
        if rtn_db[filter].empty:
            tableOccupyByOrderNumber += [" "]
            OrderTotalPax += [" "]
        else:
            orderNumber = str(rtn_db[filter]["订单号"].values[0])
            tableOccupyByOrderNumber += [orderNumber]

            totalPax = int(rtn_db[filter]["载客量"].values[0])
            OrderTotalPax += [totalPax]
    
    return_df = pd.DataFrame({
        "桌名" : tableNameArr,
        "桌台载客量" : tableOccupancy,
        "被占据的订单号" : tableOccupyByOrderNumber,
        "订单载客量" : OrderTotalPax
    })

    unassigned_df = rtn_db.copy()
    unassigned_df = unassigned_df[unassigned_df["订单ID"].isin(unassigned_orderId)]

    return return_df, unassigned_df

def rtn_upload_database(df, db_sheetname, google_auth, rtn_database_url, fernet_key, rtn_constants_dict, slice=True, rtn_control_url=None, isDb=True):
    if isDb:
        sheetname = rtn_constants_dict[db_sheetname+"_sheetname"]
        rtn_database_url = fernet_decrypt(rtn_database_url, fernet_key)

        try:
            db_sheet = google_auth.open_by_url(rtn_database_url)
            sheetname_index = db_sheet.worksheet(property="title", value=sheetname).index
            
            if slice:
                db_df = db_sheet[sheetname_index].get_as_df()
                upload_df = pd.concat([db_df, df], ignore_index=True)
            else:
                upload_df = df
                db_sheet[sheetname_index].clear()

            upload_df = rtn_datetime_column_convert(upload_df)

            if db_sheetname == "rtn_db":
                upload_df.sort_values(by="订单创建时间", ascending=False, inplace=True, ignore_index=True)

            elif db_sheetname == "food_db":
                upload_df.sort_values(by="订单ID", ascending=False, inplace=True, ignore_index=True)

            elif db_sheetname == "sm_bd":
                upload_df.sort_values(by="订单ID", ascending=False, inplace=True, ignore_index=True)

            elif db_sheetname == "payment":
                upload_df.sort_values(by="订单ID", ascending=False, inplace=True, ignore_index=True)
            
            upload_df = rtn_datetime_strftime(upload_df)

            db_sheet[sheetname_index].set_dataframe(upload_df, start="A1", nan="")

            upload_success = True
            print("上传成功")
        
        except Exception as e:
            print("上传失败, 错误描述如下: ")
            print(e)
            print()
            upload_success = False

    else:
        sheetname = rtn_constants_dict[db_sheetname]
        rtn_control_url = fernet_decrypt(rtn_control_url, fernet_key)

        try:
            control_sheet = google_auth.open_by_url(rtn_control_url)

            sheetname_index = control_sheet.worksheet(property="title", value=sheetname).index

            df = rtn_datetime_column_convert(df)


            df.sort_values(by="创建时间", ascending=False, inplace=True, ignore_index=True)

            
            df = rtn_datetime_strftime(df)
            
            control_sheet[sheetname_index].set_dataframe(df, start="A1", nan="")

            upload_success = True
            print("上传成功")
        
        except Exception as e:
            print("上传失败, 错误描述如下: ")
            print(e)
            print()
            upload_success = False


    return upload_success

def rtn_order_chit(rtn_constants_dict, other_controls, orderId, db, songTi, logoImagePath):
    with tqdm(total=100) as pbar:
        pbar.set_description("处理订单信息...")
        rtn_db = db["rtn_db"].copy()
        food_db = db["food_db"].copy()
        sm_bd = db["sm_bd"].copy()
        payment = db["payment"].copy()

        rtn_db["订单ID"] = rtn_db["订单ID"].astype(str)
        food_db["订单ID"] = food_db["订单ID"].astype(str)
        food_db["菜品ID"] = food_db["菜品ID"].astype(str)
        sm_bd["订单ID"] = sm_bd["订单ID"].astype(str)
        payment["订单ID"] = payment["订单ID"].astype(str)

        rtn_db["备注"] = rtn_db["备注"].astype(str)
        rtn_db["订单创建时间"] = pd.to_datetime(rtn_db["订单创建时间"])
        rtn_db["预订时间"] = pd.to_datetime(rtn_db["预订时间"])

        food_db["数量"] = food_db["数量"].astype(int)
        food_db["±价"] = food_db["±价"].astype(float)
        food_db["折扣"] = food_db["折扣"].astype(float)
        food_db["税前价格"] = food_db["税前价格"].astype(float)
        food_db["服务费"] = food_db["服务费"].astype(float)
        food_db["GST"] = food_db["GST"].astype(float)
        food_db["税后价格"] = food_db["税后价格"].astype(float)

        rtn_db = rtn_db[rtn_db["订单ID"] == orderId]
        food_db = food_db[food_db["订单ID"] == orderId]
        sm_bd = sm_bd[sm_bd["订单ID"] == orderId]
        payment = payment[payment["订单ID"] == orderId]

        orderNumber = str(rtn_db["订单号"].values[0])
        customerName = str(rtn_db["姓名"].values[0])
        customerPhone = str(rtn_db["电话"].values[0])
        orderDateCreated = pd.to_datetime(rtn_db["订单创建时间"].values[0]).strftime("%d/%m/%Y %H:%M")
        reserveDatetime_dt = pd.to_datetime(rtn_db["预订时间"].values[0])
        reserveDatetime = reserveDatetime_dt.strftime("%d/%m/%Y %H:%M")
        orderAttribute = str(rtn_db["订单属性"].values[0])
        orderRound = str(rtn_db["轮数"].values[0])
        pax = "{}大{}小{}幼".format(int(rtn_db["成人人数"].values[0]), int(rtn_db["儿童人数"].values[0]), int(rtn_db["幼儿人数"].values[0])) 
        orderRemark = str(rtn_db["备注"].values[0])

        svc_multiplier = float(rtn_constants_dict["service_charge_multiplier"])
        gst_multiplier = float(rtn_constants_dict["gst_multiplier"])

        svc_rate = (svc_multiplier - 1)*100
        gst_rate = (gst_multiplier - 1)*100

        outlet_address = str(rtn_constants_dict["outlet_address"])
        outlet_phone = str(rtn_constants_dict["outlet_telephone"])
        website = str(rtn_constants_dict["website"])

        display_food_db, food_db, sm_bd_df, sm_bd, payment_info = rtn_food_order_parser(food_db=food_db, sm_bd=sm_bd, payment=payment, other_controls=other_controls)
        pbar.update(12)

        pbar.set_description("PDF任务开始...")
        Document = borb_Document()
        Page = borb_Page(width=Decimal(595), height=Decimal(842))
        Document.add_page(Page)

        layout: borb_PageLayout = borb_SCL(Page)

        table0 = borb_Table(number_of_rows = 1, number_of_columns = 2)

        table0.add(borb_Image(logoImagePath, width=Decimal(80), height=Decimal(20), horizontal_alignment=borb_align.CENTERED,))
        table0.add(borb_Paragraph("订单凭证 Order Chit", font=songTi, horizontal_alignment=borb_align.LEFT,))

        table0.set_padding_on_all_cells(Decimal(1), Decimal(1), Decimal(1), Decimal(1))
        table0.no_borders()
        layout.add(table0)
        pbar.update(12)

        if orderAttribute == "堂食":
            table1 = borb_Table(number_of_rows = 4 , number_of_columns=4)
            sequence1 = [
                "订单号Order No:",
                orderNumber,
                "姓名Name:",
                customerName,
                
                "创建时间Order Placed:",
                orderDateCreated,
                "预订时间Reserved On:",
                reserveDatetime,

                "电话Phone:",
                customerPhone,
                "属性Eat-in/To-go:",
                orderAttribute,
                
                "轮数Round:",
                orderRound,
                "人数Pax:",
                pax,]

        else:
            table1 = borb_Table(number_of_rows = 3 , number_of_columns=4)
            sequence1 = [
                "订单号Order No:",
                orderNumber,
                "姓名Name:",
                customerName,
                
                "创建时间Order Placed:",
                orderDateCreated,        
                "预订时间Reserved On:",
                reserveDatetime,
                
                "电话Phone:",
                customerPhone,
                "属性Eat-in/To-go:",
                orderAttribute,]

        for i in range(len(sequence1)):
            if i % 2 == 0:
                table1.add(borb_Paragraph(str(sequence1[i]), font=songTi, horizontal_alignment=borb_align.RIGHT))
            else:
                table1.add(borb_Paragraph(str(sequence1[i]), font=songTi, horizontal_alignment=borb_align.CENTERED))
                    
        table1.set_padding_on_all_cells(Decimal(1), Decimal(1), Decimal(1), Decimal(1))
        table1.no_borders()
        layout.add(table1)
        pbar.update(12)

        if len(orderRemark) > 0:
            if orderRemark != "-":
                table2 = borb_Table(number_of_rows = 1, number_of_columns = 2)

                table2.add(borb_Paragraph("订单备注Remark:", font=songTi, horizontal_alignment=borb_align.CENTERED))
                table2.add(borb_Paragraph(orderRemark, font=songTi, horizontal_alignment=borb_align.LEFT))

                table2.set_padding_on_all_cells(Decimal(1), Decimal(1), Decimal(1), Decimal(1))
                table2.no_borders()
                layout.add(table2)
            else:
                pass
        else:
            pass
        pbar.update(12)

        if isinstance(display_food_db, pd.DataFrame):
            display_food_db["点餐INDEX"] = display_food_db["点餐INDEX"].astype(int)
            display_food_db.sort_values(by="点餐INDEX", ascending=True, inplace=True, ignore_index=True)
            
            display_food_db.drop(["点餐ID", "订单ID","菜品ID", "套餐?", "换菜?", "点餐INDEX", "服务费", "GST", "税后价格"], axis=1, inplace=True)
            display_food_db = display_food_db[["数量", "菜名/套餐名", "菜品属性", "±价", "折扣", "税前价格", "备注"]]

            table3_columns = ["量QTY", "食FOOD", "属ATTR", "±", "折DIS", "税前ST", "注RE"]

            table3 = borb_Table(number_of_rows=len(display_food_db)+1, number_of_columns=len(table3_columns))

            for h in table3_columns:
                table3.add(borb_TableCell(borb_Paragraph(h, font=songTi, horizontal_alignment=borb_align.CENTERED)))

            for index in range(len(display_food_db)):
                for column in range(len(display_food_db.columns)):
                    table3.add(borb_Paragraph(str(display_food_db.iloc[index, column]), font=songTi, horizontal_alignment=borb_align.CENTERED))

            table3.set_padding_on_all_cells(Decimal(1.5), Decimal(1.5), Decimal(1.5), Decimal(1.5))
            layout.add(borb_Paragraph("订餐详情 Food Details:", font=songTi, horizontal_alignment=borb_align.CENTERED))
            layout.add(table3)
            
        else:
            pass
        pbar.update(12)

        if isinstance(sm_bd_df, pd.DataFrame):
            if sm_bd_df.empty:
                pass
            else:
                sm_bd_df.drop(["点餐ID"], axis=1, inplace=True)
                sm_bd_df["数量"] = sm_bd_df["数量"].astype(int)
                sm_bd_df = sm_bd_df[["数量", "套餐名", "菜肴Index", "菜肴名", "备注"]]
                
                table4_columns = ["量QTY", "套SET", "#", "食FOOD", "注RE"]

                table4 = borb_Table(number_of_rows=len(sm_bd_df)+1, number_of_columns=len(table4_columns))

                for h in table4_columns:
                    table4.add(borb_TableCell(borb_Paragraph(h, font=songTi, horizontal_alignment=borb_align.CENTERED)))

                for index in range(len(sm_bd_df)):
                    for column in range(len(sm_bd_df.columns)):
                        table4.add(borb_Paragraph(str(sm_bd_df.iloc[index, column]), font=songTi, horizontal_alignment=borb_align.CENTERED))

                table4.set_padding_on_all_cells(Decimal(1.5), Decimal(1.5), Decimal(1.5), Decimal(1.5))
                layout.add(borb_Paragraph("套餐详情 Set Item Details:", font=songTi, horizontal_alignment=borb_align.CENTERED))
                layout.add(table4)
        else:
            pass
        pbar.update(12)

        if float(format(payment_info["总税后价格"], ".2f")) != 0:
            if float(format(payment_info["已付金额"], ".2f")) >= float(format(payment_info["总税后价格"], ".2f")):
                table5 = borb_Table(number_of_rows = 5, number_of_columns = 2)

                sequence5 = [
                    "SUBTOTAL:",
                    "$"+format(payment_info["总税前价格"], ".2f"),

                    "S/C ({}%):".format(int(svc_rate)),
                    "$"+format(payment_info["总服务费"], ".2f"),

                    "GST ({}%):".format(int(gst_rate)),
                    "$"+format(payment_info["总GST"], ".2f"),

                    "TOTAL:",
                    "$"+format(payment_info["总税后价格"], ".2f"),

                    "PAID:",
                    "$"+format(payment_info["已付金额"], ".2f"),]

            else:
                table5 = borb_Table(number_of_rows=1, number_of_columns=2)

                sequence5 = [
                    "DEPOSITED:",
                    "$"+format(payment_info["已付金额"], ".2f")]

        else:
            table5 = borb_Table(number_of_rows=1, number_of_columns=1)

            sequence5 = ["TOTAL: FREE OF CHARGE"]

        if len(sequence5) == 1:
            table5.add(borb_Paragraph(sequence5[0], font=songTi, horizontal_alignment=borb_align.CENTERED))
        
        else:
            for i in range(len(sequence5)):
                if i % 2 == 0:
                    table5.add(borb_Paragraph(sequence5[i], font=songTi, horizontal_alignment=borb_align.RIGHT))
                else:
                    table5.add(borb_Paragraph(sequence5[i], font=songTi, horizontal_alignment=borb_align.CENTERED))
                
        table5.set_padding_on_all_cells(Decimal(1.5), Decimal(1.5), Decimal(1.5), Decimal(1.5))
        table5.no_borders()
        layout.add(borb_Paragraph("付款详情(新加坡元) Payment Details(Singapore Dollars):", font=songTi, horizontal_alignment=borb_align.CENTERED))
        layout.add(table5)
        pbar.update(24)

        layout.add(borb_Paragraph("This chit is not an official receipt, price will be accurate on the day of spending.", font=songTi, horizontal_alignment=borb_align.CENTERED))
        layout.add(borb_Paragraph("此凭证非正式收据,实际金额以当天消费为准。", font=songTi, horizontal_alignment=borb_align.CENTERED))
        layout.add(borb_Paragraph("_______________________________________________________________________________", horizontal_alignment=borb_align.CENTERED, font=songTi))
        layout.add(borb_Paragraph("三人行|福建四川|精聚一堂 Savor the Flavors of Fujian & Sichuan", font=songTi, horizontal_alignment=borb_align.CENTERED))
        layout.add(borb_Paragraph("{} {} {}".format(outlet_address, outlet_phone, website), font=songTi, horizontal_alignment=borb_align.CENTERED))
        pbar.update(4)

        fileName = "{}_{}_{}.pdf".format(orderNumber, orderRound, reserveDatetime_dt.strftime("%Y_%m_%d_%H_%M"))  
        orderChit_subFolderName = str(rtn_constants_dict["orderChit_subFolderName"])
        export_folderName = str(rtn_constants_dict["export_folderName"])

        with open("{}/{}/{}/{}".format(os.getcwd(), export_folderName, orderChit_subFolderName, fileName), "wb") as pdf_file_handle:
            borb_PDF.dumps(pdf_file_handle, Document)  
        

        pbar.set_description("任务完成")
        print("生成完成, 文件路径在'{}'。".format("{}/{}/{}/{}".format(os.getcwd(), export_folderName, orderChit_subFolderName, fileName)))
        pbar.update(4)

def rtn_trio_list(rtn_constants_dict, other_controls, orderId, db, songTi, selected_foodOrderId):
    with tqdm(total=100) as pbar:
        pbar.set_description("处理订单信息...")
        rtn_db = db["rtn_db"].copy()
        food_db = db["food_db"].copy()
        sm_bd = db["sm_bd"].copy()
        payment = db["payment"].copy()

        rtn_db["订单ID"] = rtn_db["订单ID"].astype(str)
        food_db["订单ID"] = food_db["订单ID"].astype(str)
        food_db["菜品ID"] = food_db["菜品ID"].astype(str)
        sm_bd["订单ID"] = sm_bd["订单ID"].astype(str)
        payment["订单ID"] = payment["订单ID"].astype(str)

        rtn_db["备注"] = rtn_db["备注"].astype(str)
        rtn_db["订单创建时间"] = pd.to_datetime(rtn_db["订单创建时间"])
        rtn_db["预订时间"] = pd.to_datetime(rtn_db["预订时间"])

        food_db["数量"] = food_db["数量"].astype(int)
        food_db["±价"] = food_db["±价"].astype(float)
        food_db["折扣"] = food_db["折扣"].astype(float)
        food_db["税前价格"] = food_db["税前价格"].astype(float)
        food_db["服务费"] = food_db["服务费"].astype(float)
        food_db["GST"] = food_db["GST"].astype(float)
        food_db["税后价格"] = food_db["税后价格"].astype(float)

        rtn_db = rtn_db[rtn_db["订单ID"] == orderId]
        food_db = food_db[food_db["订单ID"] == orderId]
        sm_bd = sm_bd[sm_bd["订单ID"] == orderId]
        payment = payment[payment["订单ID"] == orderId]
        
        orderNumber = str(rtn_db["订单号"].values[0])
        customerName = str(rtn_db["姓名"].values[0])
        customerPhone = str(rtn_db["电话"].values[0])
        orderRound = str(rtn_db["轮数"].values[0])
        orderAttribute = str(rtn_db["订单属性"].values[0])
        tableAssign = str(rtn_db["桌台"].values[0])
        pax = "{}大{}小{}幼".format(int(rtn_db["成人人数"].values[0]), int(rtn_db["儿童人数"].values[0]), int(rtn_db["幼儿人数"].values[0])) 

        display_food_db, food_db, sm_bd_df, sm_bd, payment_info = rtn_food_order_parser(food_db=food_db, sm_bd=sm_bd, payment=payment, other_controls=other_controls)

        if float(payment_info["需付金额"]) == 0:
            payment_status = "全付清"
        elif float(payment_info["需付金额"]) < 0:
            payment_status = "付多"
        else:
            if float(payment_info["需付金额"]) == float(payment_info["总税后价格"]):
                payment_status = "未付"
            else:
                payment_status = "未付清"
        
        pbar.update(20)
        pbar.set_description("PDF任务开始...")

        Document = borb_Document()
        Page = borb_Page(width=Decimal(842), height=Decimal(595))
        Document.add_page(Page)

        layout: borb_PageLayout = borb_MCL(Page, number_of_columns=3)

        r_left = 28
        if orderAttribute == "堂食":
            table1 = borb_Table(number_of_rows=4 , number_of_columns=2)
            sequence1 = ["{}(单号)".format(orderNumber),
                        "{}(桌)".format(tableAssign),
                        
                        "{}(名)".format(customerName),
                        "{}(号)".format(customerPhone),

                        "{}(轮)".format(orderRound),
                        "{}(人)".format(pax),
                        
                        "支${}({})".format(format(payment_info["已付金额"], ".2f"), payment_status),
                        "桌单",]
            table1_rows = 4
            r_left -= table1_rows

        else:
            table1 = borb_Table(number_of_rows=3 , number_of_columns=2)
            sequence1 = ["{}(单号)".format(orderNumber),
                        "{}(名)".format(customerName),

                        "{}(号)".format(customerPhone),
                        "支${}({})".format(format(payment_info["已付金额"], ".2f"), payment_status),
                        
                        "桌单",
                        "",]

            table1_rows = 3
            r_left -= table1_rows

        for i in range(len(sequence1)):
            table1.add(borb_Paragraph(str(sequence1[i]), font=songTi, horizontal_alignment=borb_align.LEFT))

        table1.set_padding_on_all_cells(Decimal(1), Decimal(1), Decimal(1), Decimal(1))
        table1.no_borders()
        layout.add(table1)
        pbar.update(20)

        if isinstance(display_food_db, pd.DataFrame):
            if display_food_db.empty:
                sm_bd_df = None
                slice_dfd = None
                selectedIsSet = None
            else:
                slice_dfd = display_food_db.copy()
                slice_dfd = slice_dfd[slice_dfd["点餐ID"] != selected_foodOrderId]
                slice_dfd["点餐INDEX"] = slice_dfd["点餐INDEX"].astype(int)
                slice_dfd.sort_values(by="点餐INDEX", ascending=True, inplace=True, ignore_index=True)
                slice_dfd.drop(["点餐ID", "订单ID", "菜品ID", "点餐INDEX", "套餐?", "换菜?", "菜品属性", "±价", "折扣", "税前价格", "服务费", "GST", "税后价格"], axis=1, inplace=True)
                slice_dfd = slice_dfd[["数量", "菜名/套餐名", "备注"]]
                slice_dfd["数量"] = slice_dfd["数量"].astype(int)

                selected_food = display_food_db.copy()
                selectedIsSet = int(selected_food[selected_food["点餐ID"] == selected_foodOrderId]["套餐?"].values[0])

                if selectedIsSet == 1:
                    if isinstance(sm_bd_df, pd.DataFrame):
                        if sm_bd_df.empty:
                            sm_bd_df = None
                        else:
                            sm_bd_df = sm_bd_df[sm_bd_df["点餐ID"] == selected_foodOrderId]
                            sm_bd_df["菜肴Index"] = sm_bd_df["菜肴Index"].astype(int)
                            sm_bd_df.sort_values(by="菜肴Index", ascending=True, inplace=True, ignore_index=True)
                            foodName = str(sm_bd_df["套餐名"].values[0])
                            sm_bd_df.drop(["点餐ID", "菜肴Index", "套餐名"], axis=1, inplace=True)
                            sm_bd_df = sm_bd_df[["数量", "菜肴名", "备注"]]
                            sm_bd_df["数量"] = sm_bd_df["数量"].astype(int)
                    else:
                        sm_bd_df = None
                else:
                    foodQty = int(selected_food[selected_food["点餐ID"] == selected_foodOrderId]["数量"].values[0])
                    foodName = str(selected_food[selected_food["点餐ID"] == selected_foodOrderId]["菜名/套餐名"].values[0])
                    foodRemark = str(selected_food[selected_food["点餐ID"] == selected_foodOrderId]["备注"].values[0])

                    sm_bd_df = None
                            
        else:
            sm_bd_df = None
            slice_dfd = None 

        t2_number_of_rows = 0
        if isinstance(selectedIsSet, int):
            if int(selectedIsSet) == 1:
                if isinstance(sm_bd_df, pd.DataFrame):
                    t2_number_of_rows += len(sm_bd_df)*2+1
                else:
                    pass
            else:
                t2_number_of_rows += 1

            if slice_dfd.empty:
                pass
            else:
                t2_number_of_rows += len(slice_dfd)*2
            
        if t2_number_of_rows > 0:
            if t2_number_of_rows < r_left:
                t2_number_of_rows = r_left

            table2 = borb_Table(number_of_rows=t2_number_of_rows, number_of_columns=1)
        else:
            pass

        if isinstance(selectedIsSet, int):
            if int(selectedIsSet) == 1:
                if isinstance(sm_bd_df, pd.DataFrame):
                    table2.add(borb_Paragraph("{}".format(foodName), font=songTi, horizontal_alignment=borb_align.CENTERED))
                    r_left -= 1

                    for index in range(len(sm_bd_df)):
                        if r_left > 0:
                            table2.add(borb_Paragraph("{} {}".format(sm_bd_df.iloc[index, 0], sm_bd_df.iloc[index, 1]), font=songTi, horizontal_alignment=borb_align.LEFT))
                            r_left -= 1
                        else:
                            pass
                        
                        if r_left > 0:
                            if len(str(sm_bd_df.iloc[index, 2])) > 0:
                                if str(sm_bd_df.iloc[index, 2]) == "-":
                                    table2.add(borb_Paragraph(" ", font=songTi, horizontal_alignment=borb_align.LEFT))
                                else:
                                    table2.add(borb_Paragraph("{}".format(sm_bd_df.iloc[index, 2]), font=songTi, horizontal_alignment=borb_align.LEFT))
                            else:
                                table2.add(borb_Paragraph(" ", font=songTi, horizontal_alignment=borb_align.LEFT))
                            
                            r_left -= 1
                        else:
                            pass
                else:
                    pass
            else:
                table2.add(borb_Paragraph("{} {}".format(foodQty, foodName), font=songTi, horizontal_alignment=borb_align.LEFT))
                r_left -= 1

                if len(str(foodRemark)) > 0:
                    if str(foodRemark) == "-":
                        table2.add(borb_Paragraph(" ", font=songTi, horizontal_alignment=borb_align.LEFT))
                    else:
                        table2.add(borb_Paragraph("{}".format(foodRemark), font=songTi, horizontal_alignment=borb_align.LEFT))
                else:
                    table2.add(borb_Paragraph(" ", font=songTi, horizontal_alignment=borb_align.LEFT))
                
                r_left -= 1

            if slice_dfd.empty:
                pass
            else:
                for index in range(len(slice_dfd)):
                    if r_left > 0:
                        table2.add(borb_Paragraph("{} {}".format(slice_dfd.iloc[index, 0], slice_dfd.iloc[index, 1]), font=songTi, horizontal_alignment=borb_align.LEFT))

                        r_left -= 1
                    else:
                        pass
                    
                    if r_left > 0:
                        if len(str(slice_dfd.iloc[index, 2])) > 0:
                            if str(slice_dfd.iloc[index, 2]) == "-":
                                table2.add(borb_Paragraph(" ", font=songTi, horizontal_alignment=borb_align.LEFT))
                            else:
                                table2.add(borb_Paragraph("{}".format(slice_dfd.iloc[index, 2]), font=songTi, horizontal_alignment=borb_align.LEFT))
                        else:
                            table2.add(borb_Paragraph(" ", font=songTi, horizontal_alignment=borb_align.LEFT))
                        
                        r_left -= 1
                    else:
                        pass

        if r_left > 0:
            for _ in range(r_left):
                table2.add(borb_Paragraph(" ", font=songTi, horizontal_alignment=borb_align.LEFT))

        table2.set_padding_on_all_cells(Decimal(1), Decimal(1), Decimal(1), Decimal(1))
        table2.no_borders()

        layout.add(table2)
        pbar.update(30)

        if orderAttribute == "堂食":
            table1 = borb_Table(number_of_rows=4 , number_of_columns=2)
            sequence1 = ["{}(单号)".format(orderNumber),
                        "{}(桌)".format(tableAssign),
                        
                        "{}(名)".format(customerName),
                        "{}(号)".format(customerPhone),

                        "{}(轮)".format(orderRound),
                        "{}(人)".format(pax),
                        
                        "支${}({})".format(format(payment_info["已付金额"], ".2f"), payment_status),
                        "收银",]

        else:
            table1 = borb_Table(number_of_rows=3 , number_of_columns=2)
            sequence1 = ["{}(单号)".format(orderNumber),
                        "{}(名)".format(customerName),

                        "{}(号)".format(customerPhone),
                        "支${}({})".format(format(payment_info["已付金额"], ".2f"), payment_status),
                        
                        "收银",
                        "",]

        for i in range(len(sequence1)):
            table1.add(borb_Paragraph(str(sequence1[i]), font=songTi, horizontal_alignment=borb_align.LEFT))

        table1.set_padding_on_all_cells(Decimal(1), Decimal(1), Decimal(1), Decimal(1))
        table1.no_borders()
        layout.add(table1)
        layout.add(table2)
        pbar.update(10)

        if orderAttribute == "堂食":
            table1 = borb_Table(number_of_rows=4 , number_of_columns=2)
            sequence1 = ["{}(单号)".format(orderNumber),
                        "{}(桌)".format(tableAssign),
                        
                        "{}(名)".format(customerName),
                        "{}(号)".format(customerPhone),

                        "{}(轮)".format(orderRound),
                        "{}(人)".format(pax),
                        
                        "支${}({})".format(format(payment_info["已付金额"], ".2f"), payment_status),
                        "厨房",]

        else:
            table1 = borb_Table(number_of_rows=3 , number_of_columns=2)
            sequence1 = ["{}(单号)".format(orderNumber),
                        "{}(名)".format(customerName),

                        "{}(号)".format(customerPhone),
                        "支${}({})".format(format(payment_info["已付金额"], ".2f"), payment_status),
                        
                        "厨房",
                        "",]

        for i in range(len(sequence1)):
            table1.add(borb_Paragraph(str(sequence1[i]), font=songTi, horizontal_alignment=borb_align.LEFT))

        table1.set_padding_on_all_cells(Decimal(1), Decimal(1), Decimal(1), Decimal(1))
        table1.no_borders()
        layout.add(table1)
        layout.add(table2)
        pbar.update(15)

        fileName = "{}_{}_{}.pdf".format(orderNumber, orderRound, tableAssign)  
        trioList_subFolderName = str(rtn_constants_dict["trioList_subFolderName"])
        export_folderName = str(rtn_constants_dict["export_folderName"])

        with open("{}/{}/{}/{}".format(os.getcwd(), export_folderName, trioList_subFolderName, fileName), "wb") as pdf_file_handle:
            borb_PDF.dumps(pdf_file_handle, Document)

        pbar.set_description("任务完成")
        print("生成完成, 文件路径在'{}".format("{}/{}/{}/{}".format(os.getcwd(), export_folderName, trioList_subFolderName, fileName)))
        pbar.update(5)

def rtn_whatsapp_sender(startDate, endDate, db, outlet):
    res = pyfiglet.figlet_format("WhatsApp")
    print(res)
    time.sleep(0.15)

    from selenium import webdriver as sel_webdriver
    from selenium.webdriver.chrome.options import Options as sel_Options
    from webdriver_manager.chrome import ChromeDriverManager
    from sys import platform as sel_platform
    from selenium.webdriver.chrome.service import Service as sel_service

    rtn_db = db["rtn_db"].copy()
    rtn_db = rtn_datetime_column_convert(rtn_db)

    rtn_db["轮数"] = rtn_db["轮数"].astype(str)
    rtn_db["订单属性"] = rtn_db["订单属性"].astype(str)
    rtn_db["姓名"] = rtn_db["姓名"].astype(str)
    rtn_db["本地号码?"] = rtn_db["本地号码?"].astype(int)
    rtn_db["桌台"] = rtn_db["桌台"].astype(str)
    rtn_db["电话"] = rtn_db["电话"].astype(str)

    rtn_db["DATE"] = rtn_db["预订时间"].apply(lambda x : pd.to_datetime(dt.datetime(year=x.year, month=x.month, day=x.day)))
    dateFilter = (rtn_db["DATE"] >= startDate)&(rtn_db["DATE"] <= endDate)
    rtn_db = rtn_db[dateFilter]
    rtn_db["时间"] = rtn_db["预订时间"].apply(lambda x : x.strftime("%H:%M"))
    rtn_db = rtn_db[rtn_db["订单属性"] == "堂食"]
    rtn_db = rtn_db[rtn_db["电话"] != ""]
    rtn_db = rtn_db[rtn_db["电话"] != "-"]

    rtn_db["whatsapp"] = "尊敬的" + rtn_db["姓名"] + "您好! 这里是三人行(" + str(outlet).strip().upper() + ")中餐馆,温馨通知您在除夕的订位是" + rtn_db["轮数"] + "(" + rtn_db["时间"] + "),您的桌台号是" + rtn_db["桌台"] + "。谢谢! 祝您新年快乐! Greetings, This is San Ren Xing(" + str(outlet).strip().upper() + "), I'd like to inform you that your time slot for CNY eve is " + rtn_db["时间"] + ". Your table number is " + rtn_db["桌台"] + ". Thank you, wishing you a happy Chinese New Year! "


    options = sel_Options()

    if sel_platform == "win32":
        options.binary_location = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

    total_numbers = len(rtn_db["电话"].values)
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

    for index in range(len(rtn_db)):
        isLocalPhone = int(rtn_db["本地号码?"].values[index])
        number = str(rtn_db["电话"].values[index])
        text = str(rtn_db["whatsapp"].values[index])

        if isLocalPhone == 1:
            url = "https://web.whatsapp.com/send?phone=65" + number + "&text=" + text
        else:
            url = "https://web.whatsapp.com/send?phone=" + number + "&text" + text
        
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
                if isLocalPhone == 1:
                    url = "https://web.whatsapp.com/send?phone=65" + number + "&text=" + text
                else:
                    url = "https://web.whatsapp.com/send?phone=" + number + "&text" + text
                
                driver.get(url)
                input("按回车键继续>>>: ")
                continue
        
        elif whatsapp_select == 2:
            print("{}发送失败, 已跳过。".format(number))
            continue
    
    print("任务完成")
    driver.quit()
    return rtn_db
     
def rtn_main(google_auth, rtn_control_url, rtn_database_url, constants_sheetname):
    google_auth = google_auth
    rtn_control_url = rtn_control_url
    constants_sheetname = constants_sheetname
    rtn_database_url = rtn_database_url

    on_net = on_internet()

    if not on_net:
        print("没有网络连接。")
        print("预订程序需全程连接网络来使用。")

    else:       
        fernet_key = get_key()

        if fernet_key == 0:
            print("安全密钥错误! ")
        
        else:
            res = pyfiglet.figlet_format("Reservation")
            print(res)
            time.sleep(0.15)

            outlet = get_outlet()
            rtn_constants_dict, other_controls = rtn_get_controls(google_auth=google_auth, fernet_key=fernet_key, outlet=outlet, rtn_control_url=rtn_control_url, constants_sheetname=constants_sheetname)

            if len(rtn_constants_dict) == 0 or len(other_controls) == 0:
                pass
            else:
                exportFolderName = rtn_constants_dict["export_folderName"]
                exportSubFolderName = rtn_constants_dict["export_subFolderName"]
                orderChitSubFolderName = rtn_constants_dict["orderChit_subFolderName"]
                summary_subFolderName = rtn_constants_dict["summary_subFolderName"]
                trioList_subFolderName = rtn_constants_dict["trioList_subFolderName"]
                whatsappFile_subFolderName = rtn_constants_dict["whatsappFile_subFolderName"]

                if not os.path.exists("{}/{}/{}".format(os.getcwd(), exportFolderName, exportSubFolderName)):
                    os.makedirs("{}/{}/{}".format(os.getcwd(),exportFolderName, exportSubFolderName))
                
                if not os.path.exists("{}/{}/{}".format(os.getcwd(), exportFolderName, summary_subFolderName)):
                    os.makedirs("{}/{}/{}".format(os.getcwd(), exportFolderName, summary_subFolderName))

                if not os.path.exists("{}/{}/{}".format(os.getcwd(), exportFolderName, orderChitSubFolderName)):
                    os.makedirs("{}/{}/{}".format(os.getcwd(), exportFolderName, orderChitSubFolderName))

                if not os.path.exists("{}/{}/{}".format(os.getcwd(), exportFolderName, trioList_subFolderName)):
                    os.makedirs("{}/{}/{}".format(os.getcwd(), exportFolderName, trioList_subFolderName))

                if not os.path.exists("{}/{}/{}".format(os.getcwd(), exportFolderName, whatsappFile_subFolderName)):
                    os.makedirs("{}/{}/{}".format(os.getcwd(), exportFolderName, whatsappFile_subFolderName))

                main_menu_select = 0
                while main_menu_select != 6:
                    print("Reservation Main Menu")
                    print("预订主菜单")
                    main_menu = option_num(["预订", "汇总", "设置", "订单凭证", "三联单", "WhatsApp工具", "退出预订"])
                    time.sleep(0.25)
                    main_menu_select = option_limit(main_menu, input("在这里输入>>>: "))

                    if main_menu_select == 0:
                        reservation_menu_select = 0
                        while reservation_menu_select != 4:
                            print()
                            print()
                            reservation_menu = option_num(["新订单", "编辑订单", "筛查订单", "查看桌台状态", "返回上一菜单"])
                            time.sleep(0.25)
                            reservation_menu_select = option_limit(reservation_menu, input("在这里输入>>>: "))

                            if reservation_menu_select == 0:
                                new_order = rtn_create_order(rtn_constants_dict, other_controls, google_auth, fernet_key, rtn_database_url)
                                isUploadSuccess = rtn_upload_database(df=new_order, db_sheetname="rtn_db", google_auth=google_auth, rtn_database_url=rtn_database_url, fernet_key=fernet_key, rtn_constants_dict=rtn_constants_dict, slice=False)

                            elif reservation_menu_select == 1:
                                print("读取中...请稍等...")
                                db = rtn_fetch_database(google_auth=google_auth, fernet_key=fernet_key, rtn_database_url=rtn_database_url, rtn_constants_dict=rtn_constants_dict)
                                print("读取完成...")
                                print()
                                print()
                                print()
                                edit_order_menu = option_num(["修改订单", "编辑点餐", "编辑财务", "返回上一菜单"])
                                time.sleep(0.25)
                                edit_order_select = option_limit(edit_order_menu, input("在这里输入>>>: "))

                                if edit_order_select == 0:
                                    
                                    slice_rtn_db, delete_orderId = rtn_edit_order(rtn_constants_dict=rtn_constants_dict, other_controls=other_controls, google_auth=google_auth, fernet_key=fernet_key, rtn_database_url=rtn_database_url)

                                    if isinstance(slice_rtn_db, pd.DataFrame):
                                        if isinstance(delete_orderId, int):
                                            if delete_orderId == -404:
                                                orderId = str(slice_rtn_db["订单ID"].values[0])

                                                rtn_db = db["rtn_db"].copy()

                                                rtn_db["订单ID"] = rtn_db["订单ID"].astype(str)
                                                rtn_db = rtn_db[rtn_db["订单ID"] != orderId]

                                                upload_db = pd.concat([rtn_db, slice_rtn_db], ignore_index=True)

                                                isUploadSuccess = rtn_upload_database(df=upload_db, db_sheetname="rtn_db", google_auth=google_auth, rtn_database_url=rtn_database_url, fernet_key=fernet_key, rtn_constants_dict=rtn_constants_dict, slice=False)

                                        else:
                                            rtn_db = db["rtn_db"].copy()
                                            food_db = db["food_db"].copy()
                                            sm_bd = db["sm_bd"].copy()
                                            payment = db["payment"].copy()

                                        
                                            rtn_db["订单ID"] = rtn_db["订单ID"].astype(str)
                                            food_db["订单ID"] = food_db["订单ID"].astype(str)
                                            sm_bd["订单ID"] = sm_bd["订单ID"].astype(str)
                                            payment["订单ID"] = payment["订单ID"].astype(str)

                                            delete_orderId = str(delete_orderId)

                                            rtn_db = rtn_db[rtn_db["订单ID"] != delete_orderId]
                                            food_db = food_db[food_db["订单ID"] != delete_orderId]
                                            sm_bd = sm_bd[sm_bd["订单ID"] != delete_orderId]
                                            payment = payment[payment["订单ID"] != delete_orderId]

                                            isUploadSuccess = rtn_upload_database(df=rtn_db, db_sheetname="rtn_db", google_auth=google_auth, rtn_database_url=rtn_database_url, fernet_key=fernet_key, rtn_constants_dict=rtn_constants_dict, slice=False)
                                            isUploadSuccess = rtn_upload_database(df=food_db, db_sheetname="food_db", google_auth=google_auth, rtn_database_url=rtn_database_url, fernet_key=fernet_key, rtn_constants_dict=rtn_constants_dict, slice=False)
                                            isUploadSuccess = rtn_upload_database(df=sm_bd, db_sheetname="sm_bd", google_auth=google_auth, rtn_database_url=rtn_database_url, fernet_key=fernet_key, rtn_constants_dict=rtn_constants_dict, slice=False)
                                            isUploadSuccess = rtn_upload_database(df=payment, db_sheetname="payment", google_auth=google_auth, rtn_database_url=rtn_database_url, fernet_key=fernet_key, rtn_constants_dict=rtn_constants_dict, slice=False)
                                    else:
                                        pass

                                elif edit_order_select == 1:
                                    rtn_db = db["rtn_db"].copy()
                                    slice_food_db, slice_sm_bd, orderId = rtn_edit_food_order(google_auth=google_auth, fernet_key=fernet_key, rtn_database_url=rtn_database_url, order_concat=rtn_db, rtn_constants_dict=rtn_constants_dict, other_controls=other_controls)

                                    if isinstance(slice_food_db, pd.DataFrame):
                                        orderId = str(orderId)
                                        food_db = db["food_db"].copy()
                                        food_db["订单ID"] = food_db["订单ID"].astype(str)
                                        food_db = food_db[food_db["订单ID"] != orderId]
                                        food_db = pd.concat([food_db, slice_food_db], ignore_index=True)

                                        isUploadSuccess = rtn_upload_database(df=food_db, db_sheetname="food_db", google_auth=google_auth, rtn_database_url=rtn_database_url, fernet_key=fernet_key, rtn_constants_dict=rtn_constants_dict, slice=False)

                                    else:
                                        pass
                                        
                                    if isinstance(slice_sm_bd, pd.DataFrame):
                                        sm_bd = db["sm_bd"].copy()
                                        sm_bd["订单ID"] = sm_bd["订单ID"].astype(str)

                                        sm_bd = sm_bd[sm_bd["订单ID"] != orderId]
                                        sm_bd = pd.concat([sm_bd, slice_sm_bd], ignore_index=True)

                                        isUploadSuccess = rtn_upload_database(df=sm_bd, db_sheetname="sm_bd", google_auth=google_auth, rtn_database_url=rtn_database_url, fernet_key=fernet_key, rtn_constants_dict=rtn_constants_dict, slice=False)

                                    else:
                                        pass

                                elif edit_order_select == 2:
                                    payment_update, del_orderId = rtn_edit_finance(google_auth=google_auth, fernet_key=fernet_key, rtn_database_url=rtn_database_url, order_concat=db["rtn_db"], rtn_constants_dict=rtn_constants_dict, other_controls=other_controls)

                                    if isinstance(payment_update, pd.DataFrame):
                                        if payment_update.empty:
                                            if isinstance(del_orderId, str):
                                                orderId = str(del_orderId)
                                            else:
                                                orderId = None
                                        else:
                                            orderId = str(payment_update["订单ID"].values[0])
                                        
                                        if isinstance(orderId, str):
                                            payment = db["payment"].copy()

                                            payment["订单ID"] = payment["订单ID"].astype(str)
                                            payment["付款时间"] = pd.to_datetime(payment["付款时间"])

                                            
                                            payment = payment[payment["订单ID"] != orderId]
                                            
                                            if not payment_update.empty:
                                                payment = pd.concat([payment, payment_update], ignore_index=True)

                                            payment.sort_values(by="付款时间", ascending=False, inplace=True, ignore_index=True)
                                            isUploadSuccess = rtn_upload_database(df=payment, db_sheetname="payment", google_auth=google_auth, rtn_database_url=rtn_database_url, fernet_key=fernet_key, rtn_constants_dict=rtn_constants_dict, slice=False)
                                        else:
                                            pass
                                    else:
                                        pass
                                    
                                else:
                                    pass
                                    
                            elif reservation_menu_select == 2:
                                print("读取中...请稍等...")
                                db = rtn_fetch_database(google_auth=google_auth, fernet_key=fernet_key, rtn_database_url=rtn_database_url, rtn_constants_dict=rtn_constants_dict)
                                
                                rtn_db = db["rtn_db"].copy()
                                food_db = db["food_db"].copy()
                                sm_bd = db["sm_bd"].copy()
                                payment = db["payment"].copy()

                                rtn_db["订单ID"] = rtn_db["订单ID"].astype(str)
                                food_db["订单ID"] = food_db["订单ID"].astype(str)
                                sm_bd["订单ID"] = sm_bd["订单ID"].astype(str)
                                payment["订单ID"] = payment["订单ID"].astype(str)

                                rtn_db["备注"] = rtn_db["备注"].astype(str)
                                rtn_db["预订时间"] = pd.to_datetime(rtn_db["预订时间"])
                                rtn_db["订单创建时间"] = pd.to_datetime(rtn_db["订单创建时间"])

                                print("读取完成...")
                                print()
                                print()
                                print()

                                filter_select = 0
                                while filter_select != 29:
                                    print()
                                    print()
                                    print("选择筛查方式: ")
                                    filter_option = ["按「订单号」升序排列", 
                                                     "按「订单号」降序排列", 
                                                     "按「订单创建时间」升序排列", 
                                                     "按「订单创建时间」降序排列", 
                                                     "按不同「订单属性」筛滤", 
                                                     "按「预订时间」升序排列", 
                                                     "按「预订时间」降序排列", 
                                                     "按是否是「预订除夕?」筛滤", 
                                                     "按不同「轮数」筛滤", 
                                                     "按「成人人数」升序排列", 
                                                     "按「成人人数」降序排列", 
                                                     "按「儿童人数」升序排列", 
                                                     "按「儿童人数」降序排列", 
                                                     "按「幼儿人数」升序排列", 
                                                     "按「幼儿人数」降序排列", 
                                                     "按「载客量」升序排列", 
                                                     "按「载客量」降序排列", 
                                                     "按是否排位「桌台」筛滤", 
                                                     "按含有的「备注」内容筛滤", 
                                                     "按不含有的「备注」内容筛滤", 
                                                     "按不同「订单状态」筛滤", 
                                                     "筛滤已经点餐的订单", 
                                                     "筛滤未点餐的订单", 
                                                     "筛滤已经全额付款的订单", 
                                                     "筛滤还没付全款的订单", 
                                                     "筛滤完全没付款的订单", 
                                                     "按特定「预订时间」筛滤", 
                                                     "按特定「订单创建时间」筛滤", 
                                                     "查看有换菜的订单",
                                                     "返回上一菜单",]

                                    filter_options = option_num(filter_option)
                                    time.sleep(0.25)
                                    filter_select = option_limit(filter_options, input("在这里输入>>>: "))

                                    if filter_select != 29:
                                        if filter_select not in [18, 19, 21, 22, 23, 24, 25, 26, 27, 28]:
                                            isSpecial = False
                                            if filter_option[filter_select].find("筛滤") == -1:
                                                isUnique = False

                                                option_text = filter_option[filter_select]
                                                startIndex = option_text.find("「")+1
                                                endIndex = option_text.find("」")
                                                columnName = option_text[startIndex:endIndex]

                                                asc_startIndex = option_text.find("」")+1
                                                asc_endIndex = option_text.find("序")
                                                asc_text = option_text[asc_startIndex:asc_endIndex]
                                                if asc_text.strip() == "升":
                                                    isAscending = True
                                                else:
                                                    isAscending = False
                                            
                                            else:
                                                isUnique = True

                                                option_text = filter_option[filter_select]
                                                startIndex = option_text.find("「")+1
                                                endIndex = option_text.find("」")
                                                columnName = option_text[startIndex:endIndex]

                                        else:
                                            isSpecial = True
                                        
                                        if isSpecial:
                                            if filter_select in [18, 19]:
                                                remark = rtn_input_validation(rule="remark", title="一些备注内容", space_removal=False, rtn_constants_dict=rtn_constants_dict)

                                                if filter_select == 18:
                                                    slice_rtn_db = rtn_db[rtn_db["备注"].str.contains(remark)]
                                                else:
                                                    slice_rtn_db = rtn_db[~rtn_db["备注"].str.contains(remark)]

                                            elif filter_select in [21, 22]:
                                                existingIDinFoodDb = np.unique(food_db["订单ID"].astype(str).values).tolist()
                                                IDnotInFoodDb = []
                                                for i in range(len(rtn_db)):
                                                    if str(rtn_db.iloc[i, 1]) not in existingIDinFoodDb:
                                                        IDnotInFoodDb += [str(rtn_db.iloc[i, 1])]
                                                    else:
                                                        pass
                                                
                                                if filter_select == 22:
                                                    slice_rtn_db = rtn_db[rtn_db["订单ID"].isin(IDnotInFoodDb)]
                                                else:
                                                    slice_rtn_db = rtn_db[~rtn_db["订单ID"].isin(IDnotInFoodDb)]

                                            elif filter_select in [23, 24, 25]:
                                                complete_payment = []
                                                incomplete_payment = []
                                                no_payment = []

                                                for index in range(len(rtn_db)):
                                                    orderId = str(rtn_db.iloc[index, 1])

                                                    slice_rtn_db = rtn_db[rtn_db["订单ID"] == orderId]
                                                    slice_food_db = food_db[food_db["订单ID"] == orderId]
                                                    slice_sm_bd = sm_bd[sm_bd["订单ID"] == orderId]
                                                    slice_payment = payment[payment["订单ID"] == orderId]
                                                    
                                                    a,b,c,d, slice_payment_info = rtn_food_order_parser(food_db=slice_food_db, sm_bd=slice_sm_bd, payment=slice_payment, other_controls=other_controls)

                                                    if slice_payment_info["总税后价格"] != 0:
                                                        if slice_payment_info["已付金额"] >= slice_payment_info["总税后价格"]:
                                                            complete_payment += [orderId]
                                                        else:
                                                            if slice_payment_info["已付金额"] == 0:
                                                                no_payment += [orderId]
                                                            else:
                                                                incomplete_payment += [orderId]
                                                    else:
                                                        complete_payment += [orderId]
                                                
                                                if filter_select == 23:
                                                    slice_rtn_db = rtn_db[rtn_db["订单ID"].isin(complete_payment)]
                                                
                                                elif filter_select == 24:
                                                    slice_rtn_db = rtn_db[rtn_db["订单ID"].isin(incomplete_payment)]
                                                
                                                else:
                                                    slice_rtn_db = rtn_db[rtn_db["订单ID"].isin(no_payment)]

                                            elif filter_select == 26:
                                                slice_rtn_db = rtn_db.copy()
                                                slice_rtn_db["预订时间"] = pd.to_datetime(slice_rtn_db["预订时间"])
                                                slice_rtn_db["预订日期"] = slice_rtn_db["预订时间"].apply(lambda x : pd.to_datetime(dt.datetime(year=x.year, month=x.month, day=x.day)))

                                                start_date, end_date = get_range_date(whole_month=False)

                                                filter = (slice_rtn_db["预订日期"] >= start_date) & (slice_rtn_db["预订日期"] <= end_date)
                                                slice_rtn_db = slice_rtn_db[filter]
                                                slice_rtn_db.drop("预订日期", axis=1, inplace=True)
                                            
                                            elif filter_select == 27:
                                                slice_rtn_db = rtn_db.copy()
                                                slice_rtn_db["订单创建时间"] = pd.to_datetime(slice_rtn_db["订单创建时间"])
                                                slice_rtn_db["订单创建日期"] = slice_rtn_db["订单创建时间"].apply(lambda x : pd.to_datetime(dt.datetime(year=x.year, month=x.month, day=x.day)))

                                                start_date, end_date = get_range_date(whole_month=False)

                                                filter = (slice_rtn_db["订单创建日期"] >= start_date) & (slice_rtn_db["订单创建日期"] <= end_date)
                                                slice_rtn_db = slice_rtn_db[filter]
                                                slice_rtn_db.drop("订单创建日期", axis=1, inplace=True)

                                            elif filter_select == 28:
                                                slice_rtn_db = rtn_db.copy()
                                                slice_food_db = food_db.copy()

                                                slice_food_db["换菜?"] = slice_food_db["换菜?"].astype(int)
                                                slice_food_db["订单ID"] = slice_food_db["订单ID"].astype(str)

                                                slice_food_db = slice_food_db[slice_food_db["换菜?"] == 1]
                                                slice_orderId = slice_food_db["订单ID"].values.astype(str).tolist()

                                                slice_rtn_db = slice_rtn_db[slice_rtn_db["订单ID"].isin(slice_orderId)]

                                        else:
                                            if isUnique:
                                                slice_rtn_db_unique = np.unique(rtn_db[columnName].values.astype(str)).tolist()
                                                unique_option = option_num(slice_rtn_db_unique)
                                                time.sleep(0.25)
                                                unique_option_select = option_limit(unique_option, input("在这里输入>>>: "))

                                                if columnName == "预订除夕?":
                                                    rtn_db["预订除夕?"] = rtn_db["预订除夕?"].astype(int)
                                                    slice_rtn_db = rtn_db[rtn_db[columnName] == int(slice_rtn_db_unique[unique_option_select])]
                                                else:
                                                    slice_rtn_db = rtn_db[rtn_db[columnName] == slice_rtn_db_unique[unique_option_select]]
                                            
                                            else:
                                                slice_rtn_db = rtn_db.sort_values(by=columnName, ascending=isAscending, ignore_index=True)

                                        prtdf(slice_rtn_db)
                                        print()
                                        export_option = ["导出", "返回上一菜单"]
                                        export_options = option_num(export_option)
                                        time.sleep(0.25)
                                        export_select = option_limit(export_options, input("在这里输入>>>: "))

                                        if export_select == 0:
                                            timeNow = pd.to_datetime(dt.datetime.now()).strftime("%Y_%m_%d_%H_%M_%S")
                                            fileName = "{}.xlsx".format(timeNow)

                                            exportFolderName = rtn_constants_dict["export_folderName"]
                                            exportSubFolderName = rtn_constants_dict["export_subFolderName"]

                                            slice_rtn_db.to_excel("{}/{}/{}/{}".format(os.getcwd(), exportFolderName, exportSubFolderName, fileName), header=True, index=False)
                                            print("导出成功, 文件在'{}/{}/{}/{}'。".format(os.getcwd(), exportFolderName, exportSubFolderName, fileName))
                                            print()
                                            print()

                                    else:
                                        pass

                            elif reservation_menu_select == 3:
                                print("读取中...请稍等...")
                                db = rtn_fetch_database(google_auth=google_auth, fernet_key=fernet_key, rtn_database_url=rtn_database_url, rtn_constants_dict=rtn_constants_dict)
                                print("读取完成...")

                                print()
                                print("输入预订时间: ")
                                reserveTime = rtn_edit_datetime(datetime=0)
                                slice_rtn_db = db["rtn_db"].copy()
                                slice_rtn = db["rtn_db"].copy()
                                slice_rtn["预订时间"] = pd.to_datetime(slice_rtn["预订时间"])

                                unique_round = np.unique(slice_rtn["轮数"]).astype(str).tolist()

                                unique_round_option = []
                                for item in unique_round:
                                    unique_round_option += [item]

                                unique_round_options = option_num(unique_round_option)
                                time.sleep(0.25)
                                unique_round_select = option_limit(unique_round_options, input("在这里输入>>>: "))

                                round_selected = unique_round_option[unique_round_select]

                                table_occ, unassigned_order = rtn_table_occupancy(rtn_db=slice_rtn_db, round=round_selected, reserveTime=reserveTime, other_controls=other_controls)

                                print()
                                print()
                                print("时间: {}".format(reserveTime.strftime("%Y-%m-%d %H:%M")))
                                print("轮数: {}".format(round_selected))
                                print("占桌状态: ")
                                print()
                                prtdf(table_occ)

                                print()
                                print()
                                print("未排位订单: ")
                                print()
                                print()
                                if unassigned_order.empty:
                                    print("暂无未排位订单。")
                                else:
                                    prtdf(unassigned_order.drop("订单ID", axis=1))
                                print()
                                print()
                                print()

                    elif main_menu_select == 1:
                        print("读取中...请稍等...")
                        db = rtn_fetch_database(google_auth=google_auth, fernet_key=fernet_key, rtn_database_url=rtn_database_url, rtn_constants_dict=rtn_constants_dict)
                        
                        rtn_db = db["rtn_db"].copy()
                        food_db = db["food_db"].copy()
                        sm_bd = db["sm_bd"].copy()
                        payment = db["payment"].copy()

                        rtn_db["订单ID"] = rtn_db["订单ID"].astype(str)
                        food_db["订单ID"] = food_db["订单ID"].astype(str)
                        food_db["菜品ID"] = food_db["菜品ID"].astype(str)
                        sm_bd["订单ID"] = sm_bd["订单ID"].astype(str)
                        payment["订单ID"] = payment["订单ID"].astype(str)

                        rtn_db["备注"] = rtn_db["备注"].astype(str)
                        rtn_db["订单创建时间"] = pd.to_datetime(rtn_db["订单创建时间"])
                        rtn_db["预订时间"] = pd.to_datetime(rtn_db["预订时间"])

                        food_db["数量"] = food_db["数量"].astype(int)
                        food_db["±价"] = food_db["±价"].astype(float)
                        food_db["折扣"] = food_db["折扣"].astype(float)
                        food_db["税前价格"] = food_db["税前价格"].astype(float)
                        food_db["服务费"] = food_db["服务费"].astype(float)
                        food_db["GST"] = food_db["GST"].astype(float)
                        food_db["税后价格"] = food_db["税后价格"].astype(float)


                        print("读取完成...")
                        print()
                        print()
                        print()

                        groupby_select = 0
                        while groupby_select != 5:
                            

                            groupby_option = ["按「预订时间」汇总", "按「轮数」汇总", "按「订单属性」汇总", "按「订单创建时间」汇总", "高级汇总", "返回上一菜单"]
                            groupby_options = option_num(groupby_option)
                            time.sleep(0.25)
                            groupby_select = option_limit(groupby_options, input("在这里输入>>>: "))

                            if groupby_select != 5:

                                if groupby_select in [0, 3]:
                                    isSpecial = False
                                    filterByDateRange = True
                                
                                elif groupby_select in [1,2]:
                                    isSpecial = False
                                    filterByDateRange = False

                                elif groupby_select == 4:
                                    isSpecial = True
                                    filterByDateRange = False
                                
                                if isSpecial:
                                    print("使用高级汇总可以筛选多列的特定条件来进行汇总。")
                                    dateRangeFilterDict = {"预订时间" : True,
                                                        "轮数" : False,
                                                        "订单属性" : False,
                                                        "订单创建时间": True,
                                                        "订单状态" : False,}
                                    
                                    gb_adv_option = ["预订时间", "轮数", "订单属性", "订单创建时间", "订单状态", "退出高级汇总"]
                                    options = option_num(gb_adv_option)
                                    time.sleep(0.25)
                                    gb_adv_select = option_limit(options, input("在这里输入>>>: "))

                                    if gb_adv_select != 5:

                                        end_filter = False
                                        while not end_filter:
                                            filterByDateRange = dateRangeFilterDict[gb_adv_option[gb_adv_select]]
                                            
                                            try:
                                                slice_rtn_db = slice_rtn_db
                                            except:
                                                slice_rtn_db = rtn_db.copy()
                                            
                                            if filterByDateRange:
                                                if gb_adv_option[gb_adv_select] == "预订时间":
                                                    slice_rtn_db["DATE"] = slice_rtn_db["预订时间"].apply(lambda x : pd.to_datetime(dt.datetime(year = x.year, month=x.month, day=x.day)))
                                                elif gb_adv_option[gb_adv_select] == "订单创建时间":
                                                    slice_rtn_db["DATE"] = slice_rtn_db["订单创建时间"].apply(lambda x: pd.to_datetime(dt.datetime(year = x.year, month = x.month, day=x.day)))

                                                startDate, endDate = get_range_date(whole_month=False)

                                                date_filter = (slice_rtn_db["DATE"] >= startDate) & (slice_rtn_db["DATE"] <= endDate)
                                                slice_rtn_db = slice_rtn_db[date_filter]
                                                slice_rtn_db.drop(["DATE"], axis=1, inplace=True)
                                            
                                            else:
                                                columnName = gb_adv_option[gb_adv_select]
                                                unique_option = np.unique(slice_rtn_db[columnName].values).astype(str).tolist()
                                                unique_options = option_num(unique_option)
                                                time.sleep(0.25)
                                                unique_select = option_limit(unique_options, input("在这里输入>>>: "))

                                                unique_value = unique_option[unique_select]
                                                slice_rtn_db = slice_rtn_db[slice_rtn_db[columnName] == unique_value]
                                            
                                            print()
                                            print()
                                            print("筛选后, 目前有{}个匹配订单。".format(len(slice_rtn_db)))
                                            print()
                                            print()
                                            print("选择/继续选择筛选列: ")
                                            gb_adv_option.remove(gb_adv_option[gb_adv_select])
                                            if "退出高级汇总" in gb_adv_option:
                                                gb_adv_option.remove("退出高级汇总")
                                            
                                            if "终止继续选择筛选列" not in gb_adv_option:
                                                gb_adv_option += ["终止继续选择筛选列"]

                                            options = option_num(gb_adv_option)
                                            time.sleep(0.25)
                                            gb_adv_select = option_limit(options, input("在这里输入>>>: "))

                                            if gb_adv_select == len(options)-1:
                                                end_filter = True
                                            else:
                                                end_filter = False

                                    else:
                                        groupby_select = 5

                                else:
                                    slice_rtn_db = rtn_db.copy()
                                    startIndex = groupby_option[groupby_select].find("「")+1
                                    endIndex = groupby_option[groupby_select].find("」")
                                    columnName = groupby_option[groupby_select][startIndex:endIndex]

                                    if filterByDateRange:
                                        if columnName == "预订时间":
                                            slice_rtn_db["DATE"] = slice_rtn_db["预订时间"].apply(lambda x : pd.to_datetime(dt.datetime(year = x.year, month=x.month, day=x.day)))
                                        elif columnName == "订单创建时间":
                                            slice_rtn_db["DATE"] = slice_rtn_db["订单创建时间"].apply(lambda x: pd.to_datetime(dt.datetime(year = x.year, month = x.month, day=x.day)))

                                        startDate, endDate = get_range_date(whole_month=False)

                                        date_filter = (slice_rtn_db["DATE"] >= startDate) & (slice_rtn_db["DATE"] <= endDate)
                                        slice_rtn_db = slice_rtn_db[date_filter]
                                        slice_rtn_db.drop(["DATE"], axis=1, inplace=True)
                                    
                                    else:
                                        unique_option = np.unique(slice_rtn_db[columnName].values).astype(str).tolist()
                                        unique_options = option_num(unique_option)
                                        time.sleep(0.25)
                                        unique_select = option_limit(unique_options, input("在这里输入>>>: "))

                                        unique_value = unique_option[unique_select]
                                        slice_rtn_db = slice_rtn_db[slice_rtn_db[columnName] == unique_value]

                                if groupby_select != 5:
                                    slice_orderId = slice_rtn_db["订单ID"].values.astype(str).tolist()
                                    slice_food_db = food_db.copy()
                                    slice_food_db = slice_food_db[slice_food_db["订单ID"].isin(slice_orderId)]

                                    slice_sm_bd = sm_bd.copy()
                                    slice_sm_bd = slice_sm_bd[slice_sm_bd["订单ID"].isin(slice_orderId)]

                                    slice_payment = payment.copy()
                                    slice_payment = slice_payment[slice_payment["订单ID"].isin(slice_orderId)]

                                    slice_display_food_db, slice_food_db, slice_sm_bd_df, slice_sm_bd, slice_payment_info = rtn_food_order_parser(food_db=slice_food_db, sm_bd=slice_sm_bd, payment=slice_payment, other_controls=other_controls)

                                    if slice_rtn_db.empty:
                                        pass
                                    else:
                                        slice_rtn_db["成人人数"] = slice_rtn_db["成人人数"].astype(int)
                                        slice_rtn_db["儿童人数"] = slice_rtn_db["儿童人数"].astype(int)
                                        slice_rtn_db["幼儿人数"] = slice_rtn_db["幼儿人数"].astype(int)
                                        slice_rtn_db["载客量"] = slice_rtn_db["载客量"].astype(int)

                                        print("订单汇总: ")
                                        print("订单一共{}单。".format(len(slice_rtn_db)))
                                        print("一共{}成人。".format(slice_rtn_db["成人人数"].sum()))
                                        print("一共{}儿童。".format(slice_rtn_db["儿童人数"].sum()))
                                        print("一共{}幼儿。".format(slice_rtn_db["幼儿人数"].sum()))
                                        print()
                                        print()

                                    if isinstance(slice_display_food_db, pd.DataFrame):
                                        if not slice_display_food_db.empty:
                                            slice_display_food_db["菜名/套餐名(备注)"] = slice_display_food_db["菜名/套餐名"] + "(" + slice_display_food_db["备注"] + ")"
                                            food_summary = slice_display_food_db.groupby("菜名/套餐名(备注)")[["数量", "税前价格", "服务费", "GST", "税后价格"]].sum()
                                            food_summary.reset_index(inplace=True)
                                            food_summary["数量"] = food_summary["数量"].astype(int)
                                            print("菜品汇总: ")
                                            prtdf(food_summary)

                                        else:
                                            print("暂无菜品汇总。")
                                            food_summary = -404

                                    else:
                                        print("暂无菜品汇总。")
                                        food_summary = -404
                                        
                                    print()
                                    print()
                                    if isinstance(slice_sm_bd_df, pd.DataFrame):
                                        if not slice_sm_bd_df.empty:
                                            slice_sm_bd_df["(套餐名)菜肴名(备注)"] = "(" + slice_sm_bd_df["套餐名"] + ")" + slice_sm_bd_df["菜肴名"] + "(" + slice_sm_bd_df["备注"] + ")"
                                            set_menu_summary = slice_sm_bd_df.groupby("(套餐名)菜肴名(备注)")[["数量"]].sum()
                                            set_menu_summary.reset_index(inplace=True)
                                            set_menu_summary["数量"] = set_menu_summary["数量"].astype(int)
                                            print("套餐内菜肴汇总: ")
                                            prtdf(set_menu_summary)
                                        else:
                                            print("暂无套餐内菜肴汇总。")
                                            set_menu_summary = -404
                                    else:
                                        print("暂无套餐内菜肴汇总。")
                                        set_menu_summary = -404
                                    
                                    if isinstance(food_summary, pd.DataFrame):
                                        print("财务详情汇总: ")
                                        for key, value in slice_payment_info.items():
                                            print("{} {} ".format(key, format(float(value), ".2f")))
                                    else:
                                        print("暂无财务详情汇总。")

                                    if isinstance(food_summary, pd.DataFrame):
                                        print()
                                        print()
                                        save_options = option_num(["导出", "返回上一菜单"])
                                        time.sleep(0.25)
                                        save_select = option_limit(save_options, input("在这里输入>>>: "))

                                        if save_select == 0:
                                            if isinstance(set_menu_summary, pd.DataFrame):
                                                save_set_menu = True
                                            else:
                                                save_set_menu = False
                                            
                                            timeNow = pd.to_datetime(dt.datetime.now()).strftime("%Y_%m_%d_%H_%M_%S")
                                            fileName = "{}.xlsx".format(timeNow)

                                            export_folderName = rtn_constants_dict["export_folderName"]
                                            summary_subFolderName = rtn_constants_dict["summary_subFolderName"]

                                            if save_set_menu:
                                                with pd.ExcelWriter("{}/{}/{}/{}".format(os.getcwd(), export_folderName, summary_subFolderName, fileName)) as writer:
                                                    food_summary.to_excel(writer, sheet_name="food_summary", index=False, header=True)
                                                    set_menu_summary.to_excel(writer, sheet_name="set_menu_summary", index=False, header=True)
                                            else:
                                                food_summary.to_excel("{}/{}/{}/{}".format(os.getcwd(), export_folderName, summary_subFolderName, fileName), sheet_name="food_summary", index=False, header=True)
                                            
                                            print("导出成功, 文件在'{}/{}/{}/{}'。".format(os.getcwd(), export_folderName, summary_subFolderName, fileName))
                                        else:
                                            pass
                                    else:
                                        pass
                    
                    elif main_menu_select == 2:
                        setting_menu_select = 0
                        while setting_menu_select != 2:
                            setting_menu = option_num(["菜品菜单", "桌台", "返回上一菜单"])
                            time.sleep(0.25)
                            setting_menu_select = option_limit(setting_menu, input("在这里输入>>>: "))

                            if setting_menu_select == 0:
                                foodMenu_menu_select = 0
                                while foodMenu_menu_select != 3:
                                    foodMenu_menu = option_num(["查看单点菜单", "查看套餐", "编辑菜单", "返回上一菜单"])
                                    time.sleep(0.25)
                                    foodMenu_menu_select = option_limit(foodMenu_menu, input("在这里输入>>>: "))

                                    if foodMenu_menu_select == 0:
                                        print()
                                        print()
                                        print()
                                        if other_controls["acm"].empty:
                                            print("暂无单点。")
                                        else:
                                            print("单点菜单详情: ")
                                            prtdf(other_controls["acm"].drop(["菜品ID"], axis=1))
                                        print()
                                        print()
                                        print()
                                    
                                    elif foodMenu_menu_select == 1:
                                        print()
                                        print()
                                        if other_controls["sm"].empty:
                                            print("暂无套餐。")
                                        else:
                                            print("套餐菜单详情: ")
                                            for index in range(len(other_controls["sm"])):
                                                foodName = str(other_controls["sm"].iloc[index, 1])
                                                foodPrice = format(float(other_controls["sm"].iloc[index, 2]), ".2f")
                                                food_items = str(other_controls["sm"].iloc[index, 3]).split(",")

                                                print("套餐{}. {}".format(index+1, foodName))
                                                print("价格: {}".format(foodPrice))

                                                for i in range(len(food_items)):
                                                    print("第{}道菜: {}".format(i+1, food_items[i]))
                                                
                                                print()
                                        print()
                                        print()
                                    
                                    elif foodMenu_menu_select == 2:
                                        acm, sm, other_controls = rtn_edit_menu(rtn_constants_dict=rtn_constants_dict, other_controls=other_controls)

                                        isUploadSuccess = rtn_upload_database(df=acm, db_sheetname="ala_carte_menu_sheetname", google_auth=google_auth, rtn_database_url=rtn_database_url, fernet_key=fernet_key, rtn_constants_dict=rtn_constants_dict, slice=False, rtn_control_url=rtn_control_url, isDb=False)
                                        isUploadSuccess = rtn_upload_database(df=sm, db_sheetname="set_menu_sheetname", google_auth=google_auth, rtn_database_url=rtn_database_url, fernet_key=fernet_key, rtn_constants_dict=rtn_constants_dict, slice=False, rtn_control_url=rtn_control_url, isDb=False)
                                        print()
                                        print()
                                        print()
                            
                            elif setting_menu_select == 1:
                                tableControl_menu_select = 0
                                while tableControl_menu_select != 2:
                                    tableControl_menu = option_num(["查看桌台", "编辑桌台", "返回上一菜单"])
                                    time.sleep(0.25)
                                    tableControl_menu_select = option_limit(tableControl_menu, input("在这里输入>>>: "))

                                    if tableControl_menu_select == 0:
                                        print()
                                        print()
                                        print()
                                        if other_controls["tc"].empty:
                                            print("暂无桌台。")
                                        else:
                                            prtdf(other_controls["tc"].drop("桌子ID", axis=1))

                                        print()
                                        print()
                                        print()
                                    
                                    elif tableControl_menu_select == 1:
                                        tc, other_controls = rtn_edit_tables(rtn_constants_dict=rtn_constants_dict, other_controls=other_controls)
                                        isUploadSuccess = rtn_upload_database(df=tc, db_sheetname="table_control_sheetname", google_auth=google_auth, rtn_database_url=rtn_database_url, fernet_key=fernet_key, rtn_constants_dict=rtn_constants_dict, slice=False, rtn_control_url=rtn_control_url, isDb=False)
                                        print()

                    elif main_menu_select == 3:
                        outlet = get_outlet()
                        stock_count_foldername = "{}盘点文件".format(outlet.strip().capitalize())
                        songti_filename = "SongTi.ttf"

                        export_folderName = rtn_constants_dict["export_folderName"]
                        logo_fileName = rtn_constants_dict["logo_fileName"]
                        logo_path = "{}/{}/{}".format(os.getcwd(), export_folderName, logo_fileName)

                        if not os.path.exists("{}/{}/{}".format(os.getcwd(), stock_count_foldername, songti_filename)):
                            print("宋体TTF文件'{}'不存在, 订单凭证的生成无法继续".format(songti_filename))
                            print("请把宋体TTF文件'{}'保存至盘点文件名的目录下, 具体路径需在:'{}/{}/{}'".format(songti_filename, os.getcwd(), stock_count_foldername, songti_filename))

                        else:
                            print("读取字体文件中...")
                            custom_font_path = pathlib.Path("{}/{}/{}".format(os.getcwd(), stock_count_foldername, songti_filename))
                            songTi = borb_TrueTypeFont.true_type_font_from_file(custom_font_path)
                            print("字体文件读取完成。")
                            print()
                            if not os.path.exists(logo_path):
                                print("公司标识图片'{}'不存在, 订单凭证的生成无法继续".format(logo_fileName))
                                print("请把公司标识图片'{}'保存至预订导出的目录下, 具体路径需在: {}".format(logo_fileName, logo_path))
                            else:
                                logoImagePath = pathlib.Path(logo_path)

                                print("读取订单信息, 请稍等...")
                                db = rtn_fetch_database(google_auth=google_auth, fernet_key=fernet_key, rtn_database_url=rtn_database_url, rtn_constants_dict=rtn_constants_dict)
                                print("订单信息读取完成。")
                                print()
                                print()
                                print()
                                orderId = rtn_select_order(order_concat=db["rtn_db"], rtn_constants_dict=rtn_constants_dict)

                                if isinstance(orderId, str):
                                    rtn_order_chit(rtn_constants_dict=rtn_constants_dict, other_controls=other_controls, orderId=orderId, db=db, songTi=songTi, logoImagePath=logoImagePath)

                                else:
                                    pass

                    elif main_menu_select == 4:
                        outlet = get_outlet()
                        stock_count_foldername = "{}盘点文件".format(outlet.strip().capitalize())
                        songti_filename = "SongTi.ttf"

                        if not os.path.exists("{}/{}/{}".format(os.getcwd(), stock_count_foldername, songti_filename)):
                            print("宋体TTF文件'{}'不存在, 三联单的生成无法继续".format(songti_filename))
                            print("请把宋体TTF文件'{}'保存至盘点文件名的目录下, 具体路径需在:'{}/{}/{}'".format(songti_filename, os.getcwd(), stock_count_foldername, songti_filename))
                        else:
                            print("读取字体文件中...")
                            custom_font_path = pathlib.Path("{}/{}/{}".format(os.getcwd(), stock_count_foldername, songti_filename))
                            songTi = borb_TrueTypeFont.true_type_font_from_file(custom_font_path)
                            print("字体文件读取完成。")
                            print()
                            print("读取订单信息, 请稍等...")
                            db = rtn_fetch_database(google_auth=google_auth, fernet_key=fernet_key, rtn_database_url=rtn_database_url, rtn_constants_dict=rtn_constants_dict)
                            print("订单信息读取完成。")
                            print()
                            print()
                            print()
                            trioListSelect = 0
                            while trioListSelect != 1:
                                trioListOption = option_num(["生成三联单", "返回上一菜单"])
                                time.sleep(0.25)
                                trioListSelect = option_limit(trioListOption, input("在这里输入>>>: "))

                                if trioListSelect == 0:
                                    print("选择订单: ")
                                    orderId = rtn_select_order(order_concat=db["rtn_db"], rtn_constants_dict=rtn_constants_dict)

                                    if isinstance(orderId, str):
                                        rtn_db = db["rtn_db"].copy()
                                        food_db = db["food_db"].copy()
                                        sm_bd = db["sm_bd"].copy()
                                        payment = db["payment"].copy()

                                        rtn_db["订单ID"] = rtn_db["订单ID"].astype(str)
                                        food_db["订单ID"] = food_db["订单ID"].astype(str)
                                        food_db["菜品ID"] = food_db["菜品ID"].astype(str)
                                        sm_bd["订单ID"] = sm_bd["订单ID"].astype(str)
                                        payment["订单ID"] = payment["订单ID"].astype(str)

                                        rtn_db["备注"] = rtn_db["备注"].astype(str)
                                        rtn_db["订单创建时间"] = pd.to_datetime(rtn_db["订单创建时间"])
                                        rtn_db["预订时间"] = pd.to_datetime(rtn_db["预订时间"])

                                        food_db["数量"] = food_db["数量"].astype(int)
                                        food_db["±价"] = food_db["±价"].astype(float)
                                        food_db["折扣"] = food_db["折扣"].astype(float)
                                        food_db["税前价格"] = food_db["税前价格"].astype(float)
                                        food_db["服务费"] = food_db["服务费"].astype(float)
                                        food_db["GST"] = food_db["GST"].astype(float)
                                        food_db["税后价格"] = food_db["税后价格"].astype(float)

                                        rtn_db = rtn_db[rtn_db["订单ID"] == orderId]
                                        food_db = food_db[food_db["订单ID"] == orderId]
                                        sm_bd = sm_bd[sm_bd["订单ID"] == orderId]
                                        payment = payment[payment["订单ID"] == orderId]

                                        display_food_db, food_db, sm_bd_df, sm_bd, payment_info = rtn_food_order_parser(food_db=food_db, sm_bd=sm_bd, payment=payment, other_controls=other_controls)
                                        print()
                                        print()
                                        print("选择要展示在三联单的套餐: ")
                                        print("如果一个订单点了多个套餐, 只有你选择的套餐能展示在三联单。")
                                        print("如果没有点套餐, 你也可以选单点菜品。")
                                        print()
                                        print()
                                        foodOrderId = rtn_select_foodOrderId(display_food_db=display_food_db)

                                        if isinstance(foodOrderId, str):
                                            rtn_trio_list(rtn_constants_dict=rtn_constants_dict, other_controls=other_controls, orderId=orderId, db=db, songTi=songTi, selected_foodOrderId=foodOrderId)
                                        else:
                                            pass

                                    else:
                                        pass

                    elif main_menu_select == 5:
                        print("WhatsApp批量发信息工具只能在电脑端使用。")
                        print("移动端设备无法使用。")
                        print()
                        print()
                        options = option_num(["我用电脑了, 继续", "退出WhatsApp工具"])
                        time.sleep(0.25)
                        option_select = option_limit(options, input("在这里输入: "))

                        if option_select == 0:
                            print("读取订单信息, 请稍等...")
                            db = rtn_fetch_database(google_auth=google_auth, fernet_key=fernet_key, rtn_database_url=rtn_database_url, rtn_constants_dict=rtn_constants_dict)
                            print("订单信息读取完成。")
                            print()
                            print()
                            print()
                            print("你要发给预订什么时候的订单? ")
                            print("如果仅要筛选一天, 开始和结束日期是一样的就行")
                            startDate, endDate = get_range_date(whole_month=False)
                            rtn_db_whatsapp = rtn_whatsapp_sender(startDate=startDate, endDate=endDate, db=db, outlet=outlet)

                            if isinstance(rtn_db_whatsapp, pd.DataFrame):
                                options = option_num(["保存WhatsApp处理文件", "返回上一菜单"])
                                time.sleep(0.25)
                                option_select = option_limit(options, input("在这里输入>>>: "))

                                if option_select == 0:
                                    export_folderName = rtn_constants_dict["export_folderName"]
                                    whatsappFile_subFolderName = rtn_constants_dict["whatsappFile_subFolderName"]
                                    timeNow = dt.datetime.now().strftime("%Y_%m_%d_%H_%M")
                                    fileName = "{}_Whatsapp.xlsx".format(timeNow)

                                    rtn_db_whatsapp.to_excel("{}/{}/{}/{}".format(os.getcwd(), export_folderName, whatsappFile_subFolderName, fileName), header=True, index=False)
                                    print("WhatsApp处理文件保存成功, 具体路径在: '{}'。".format("{}/{}/{}/{}".format(os.getcwd(), export_folderName, whatsappFile_subFolderName, fileName)))
                                    print()
                                    print()
                                else:
                                    print()
                                    print()
                            else:
                                print()
                                print()

def main(database_url, db_setting_url, serialized_rule_filename, service_filename, constants_sheetname, google_auth, box_num, drink_num, promo_num, lun_sales, lun_gc, tb_sales, tb_gc, lun_fwc, lun_kwc, tb_fwc, tb_kwc, night_fwc, night_kwc, script_backup_filename, script, wifi, backup_foldername,cashier_on_duty, drink_on_duty, box_on_duty, payslip_on_duty, do_not_show_menu, rtn_control_url, rtn_database_url):
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
    rtn_control_url = rtn_control_url
    rtn_database_url = rtn_database_url

    if do_not_show_menu:
        night_audit_main(database_url, db_setting_url, serialized_rule_filename, service_filename, constants_sheetname, google_auth, box_num, drink_num, promo_num, lun_sales, lun_gc, tb_sales, tb_gc, lun_fwc, lun_kwc, tb_fwc, tb_kwc, night_fwc, night_kwc, script_backup_filename, script, wifi, backup_foldername,cashier_on_duty, drink_on_duty, box_on_duty, payslip_on_duty)

    else:
        if int(np.random.randint(1,7,size=1)[0]) % 2 == 0:
            backup_script(script_backup_filename, script, backup_foldername)
        else:
            pass

        res = pyfiglet.figlet_format("San Ren Xing Super App")
        print(res)
        time.sleep(0.15)
        
        SRX_take_input = 0
        while SRX_take_input != 6:
            print()
            print("Main Menu")
            options = option_num(["关帐", "盘点", "排班", "生成酒水明细表", "预订", "工具箱", "终止Super App"])
            time.sleep(0.25)
            SRX_take_input = option_limit(options, input("在这里输入>>>: "))

            if SRX_take_input == 0:
                night_audit_main(database_url, db_setting_url, serialized_rule_filename, service_filename, constants_sheetname, google_auth, box_num, drink_num, promo_num, lun_sales, lun_gc, tb_sales, tb_gc, lun_fwc, lun_kwc, tb_fwc, tb_kwc, night_fwc, night_kwc, script_backup_filename, script, wifi, backup_foldername,cashier_on_duty, drink_on_duty, box_on_duty, payslip_on_duty)

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
                rtn_main(google_auth, rtn_control_url, rtn_database_url, constants_sheetname)

            elif SRX_take_input == 5:

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
    main(database_url, db_setting_url, serialized_rule_filename, service_filename, constants_sheetname, google_auth, box_num, drink_num, promo_num, lun_sales, lun_gc, tb_sales, tb_gc, lun_fwc, lun_kwc, tb_fwc, tb_kwc, night_fwc, night_kwc, script_backup_filename, script, wifi, backup_foldername,cashier_on_duty, drink_on_duty, box_on_duty, payslip_on_duty, do_not_show_menu, rtn_control_url, rtn_database_url)
