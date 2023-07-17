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
from requests.structures import CaseInsensitiveDict
import pygsheets
import string
import pickle
from cryptography.fernet import Fernet
import ast
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from tqdm import tqdm

'''
import requests
import urllib
from cryptography.fernet import Fernet

def on_internet():
    url = "https://www.google.com.sg/"
    timeout = 5
    try:
        request = requests.get(url, timeout=timeout)
        return True
    except (requests.ConnectionError, requests.Timeout) as exception:
        return False

githubUserName = ""
githubRepoName = ""
githubBranchName = "main"
script_backup_filename = ""

on_net = on_internet()

if not on_net:
    wifi = False

    print("警告! 没有网络连接! ")
    print("离线模式下，程序将会全程采用本地备份数据来运行")
    print("可能有一些没有及时更新的资料而导致生成的报表不准确")
    print("非常不推荐你使用离线模式关帐！")
    print()
    for i, name in enumerate(["离线模式继续运行","终止运行"]):
        print("{}扣{}".format(name, i))

    try:    
        userInputOne = int(input(": "))

        if userInputOne == 0:
            with open("fernet_key.txt", "rb") as keyfile:
                fernet_key = keyfile.read()
            
            with open(script_backup_filename, "rb") as sfile:
                script = sfile.read()
            
            fernet_handler = Fernet(fernet_key)
            script = fernet_handler.decrypt(script).decode()
            exec(script)

        elif userInputOne == 1:
            pass

        else:
            print("无效命令")

    except:
        print("无效命令")

else:
    wifi = True
    URL = "{}/{}/{}/{}/{}".format("https://raw.githubusercontent.com", githubUserName, githubRepoName, githubBranchName, "night-audit.py")
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
on_duty = ""

exec(open("Local Script.py", encoding="utf-8").read())

'''

database_url = "gAAAAABksmmj1NC7ooHFZ7wTi9d1dhk-Q6i1Y_GWWb24xVpLfC0lmNa5vsXGDhCYwRO_NLy8NDge8xV6zna6ny1pDagdp77f8eE9nrOmRWdzeEYMi21hbEEY88G0lnfJPl0XZb2kN6M7BWasrkIp3TeDgSFmdnigSvlT4NJn92lVdCKJIJXCq5Qik4qj52Q-JJU7LZ5Xc_Lf"
db_setting_url = "gAAAAABksml2tGtCXmmE6x14Fj-_f9RKeL7pMvBlrDn1puSotmBKVAZzWRaYLQq0FRhe2MHYX6JQMDzzeUKI5lKjADOYSKWMaA3EKicdBC94U6KDqjMf71prmgXaVKu-dq_W9NBJtgDKKU7HfmRyKU6mE1Di1Nge1wt4FFuHRjJqIxvMO6eX9OQHOY8mEDYDkcLfP0GKvXu2"
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

def get_dfs(google_auth, db_setting_url, service_filename, constants_sheetname, serialized_rule_filename, outlet, fernet_key, wifi):

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
            with open(serialized_rule_filename, "wb") as f:
                pickle.dump(backup_list, f)

        except Exception as e:
            print()
            print("从网络读取规则文件失败，错误描述如下: ")
            print()
            print(e)

            with open(serialized_rule_filename, "rb") as f:
                backup_list = pickle.load(f)

            k_df = backup_list[0]
            rule_df_dict = backup_list[1]

            outlet_col_index = k_df.columns.get_loc(key=outlet)

            k_dict = {}
            for index in range(len(k_df)):
                k_dict.update({ str(k_df.iloc[index, 0]) : str(k_df.iloc[index, outlet_col_index]) })

            print("已采用本地备份规则文件数据")

    else:
        with open(serialized_rule_filename, "rb") as f:
            backup_list = pickle.load(f) #a list of two objects, 0 is k_df, 1 is rule_df_dict

        k_df = backup_list[0]
        rule_df_dict = backup_list[1]

        outlet_col_index = k_df.columns.get_loc(key=outlet)

        k_dict = {}
        for index in range(len(k_df)):
            k_dict.update({ str(k_df.iloc[index, 0]) : str(k_df.iloc[index, outlet_col_index]) })

    return k_dict, rule_df_dict

def get_book_dfs(k_dict):

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
                outlet_loc = outlet
            except NameError:
                outlet_loc = "门店【待编辑】"

            print("The programme had detected that you are using a book from local machine to do closing")
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
                     "NET SALES YESTERDAY",
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

def retrieve_database(k_dict, google_auth, fernet_key, wifi, database_url):

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

                if os.path.exists("{}/{}".format(os.getcwd(), local_db_filename)):
                    local_db_df = pd.read_excel("{}/{}".format(os.getcwd(), local_db_filename), sheet_name=name)
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
                local_db_df = pd.read_excel("{}/{}".format(os.getcwd(), local_db_filename), sheet_name=name)
                take_databases.update({ name : local_db_df })

            print("已采用本地备份的数据库信息")

    else:
        take_databases = {}
        for name in sheet_names:
            local_db_df = pd.read_excel("{}/{}".format(os.getcwd(), local_db_filename), sheet_name=name)
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

def parse_tabox(k_dict, google_auth, fernet_key, box_num, rule_df_dict, take_databases, book_dict, date_dict, wifi):
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

                local_database_filename = str(k_dict["local_database_filename"])
                tabox_sheet_name = str(k_dict["takeaway_box_sheetname"])

                tabox_df = pd.read_excel(local_database_filename, sheet_name=tabox_sheet_name)

                print("已采用本地备份的打包盒出入库数据。")
                print("注意确保本地备份的打包盒名字和数据库里的是一模一样的，当天所有出入库的数据是正确的。")
                print("如果需要更改本地备份数据，请先强制停止此程序，改好后重新运行。")
                print("如果无需更改任何数据, 按任意键继续运行程序")
                input(":")
    else:
        local_database_filename = str(k_dict["local_database_filename"])
        tabox_sheet_name = str(k_dict["takeaway_box_sheetname"])

        tabox_df = pd.read_excel(local_database_filename, sheet_name=tabox_sheet_name)

        print()
        print("无网络连接，已采用本地备份的打包盒出入库数据")
        print("注意确保本地备份的打包盒名字和数据库里是一模一样的，当天所有出入库的数据是正确的。")
        print("如果需要更改本地备份数据，请先强制停止此程序，改好后重新运行。")
        print("如果无需更改任何数据, 按任意键继续运行程序")
        input(":")

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

                    criteria = (operation != "zero") and (operation != "change value")

                    if criteria:
                        key0 = str(etb_df["BOX VALUE KEY TO OPERATE ON"][op])
                        key1 = str(etb_df["WITH KEY"][op])

                        if operation == "plus equal":
                            box_value[key0] += box_value[key1]

                        elif operation == "minus equal":
                            box_value[key0] -= box_value[key1]

                        elif operation == "divide equal":
                            box_value[key0] /= box_value[key1]

                        elif operation == "times equal":
                            box_value[key0] *= box_value[key1]

                        elif operation == "equal":
                            box_value[key0] = box_value[key1]

                        else:
                            pass

                    else:
                        if operation == "zero":
                            box_value[key0] = 0

                        elif operation == "change value":
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

def parse_alert(stock_alert_bool, date_dict, unparsed_alert, understock_alert_bool, overstock_alert_bool, alert_freq, alert_type):
    dfb = date_dict["dfb"]
    understock_alert = unparsed_alert["understock_alert"]
    understock_other = unparsed_alert["understock_other"]
    stock_all = unparsed_alert["stock_all"]
    overstock_alert = unparsed_alert["overstock_alert"]

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
                message_string = "{} \n".format(title_dict["stock_all"])
                for t in range(len(stock_all)):
                    message_string += "{} \n".format(stock_all[t])

            else:
                if len(understock_alert) > 0:
                    message_string = "{} \n".format(title_dict["understock_alert"])
                    for t in range(len(understock_alert)):
                        message_string += "{} \n".format(understock_alert[t])

                    message_string += "###################\n"
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
                message_string = ""

            if len(overstock_alert) > 0:
                message_string += "###################\n"
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

                for ppl in mail_receivers:
                    print("正在发送电邮给{}".format(ppl))

                    msg["To"] = ppl
                    smtp_obj.sendmail(mail_sender, pp, msg.as_string())

                    print("{} 电邮发送成功".format(ppl))

            except smtplib.SMTPException:
                print("电邮发送失败")
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
    pay_breakdown = book_dict["pay_breakdown"]
    pay_brkdwn = book_dict["pay_brkdwn"]
    read_ = book_dict["read_"]
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

def parse_drink_stock(k_dict, fernet_key, google_auth, drink_num, value_dict, take_databases, date_dict, wifi):
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
                local_database_filename = str(k_dict["local_database_filename"])
                drink_stock_sheetname = str(k_dict["drink_stock_sheetname"])

                drink_df = pd.read_excel(local_database_filename, sheet_name=drink_stock_sheetname)

                print("已采用本地备份的酒水出入库数据。")
                print("注意确保本地备份的酒水名字和数据库里的是一模一样的，当天所有出入库的数据是正确的。")
                print("如果需要更改本地备份数据，请先强制停止此程序，改好后重新运行。")
                print("如果无需更改任何数据, 按任意键继续运行程序")
                input(":")
    else:
        local_database_filename = str(k_dict["local_database_filename"])
        drink_stock_sheetname = str(k_dict["drink_stock_sheetname"])

        drink_df = pd.read_excel(local_database_filename, sheet_name=drink_stock_sheetname)

        print("无网络连接，已采用本地备份的酒水出入库数据")
        print("注意确保本地备份的酒水名字和数据库里的是一模一样的，当天所有出入库的数据是正确的。")
        print("如果需要更改本地备份数据，请先强制停止此程序，改好后重新运行。")
        print("如果无需更改任何数据, 按任意键继续运行程序")
        input(":")

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

    local_db_filename = str(k_dict["local_database_filename"])
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
            cum_net_sales_tdy = value_dict["cmns_tdy"]
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

def upload_db(database_url, take_databases, k_dict, fernet_key, google_auth, box_num, drink_num, wifi, outlet):
    #print("正在上传数据...")

    auto_hour = 20
    auto_min = 45

    NOW = dt.datetime.now()
    AUTO_TIME = dt.datetime(NOW.year, NOW.month, NOW.day, auto_hour, auto_min)

    local_database_filename = str(k_dict["local_database_filename"])
    online_database_url = database_url
    online_database_url = fernet_decrypt(online_database_url, fernet_key)

    drink_db_sheetname = k_dict["drink_database_sheetname"]
    takeaway_box_db_sheetname = k_dict["takeaway_box_database_sheetname"]
    financial_db_sheetname = k_dict["financial_database_sheetname"]
    promotion_db_sheetname = k_dict["promotion_database_sheetname"]
    shift_db_sheetname = "{}_shift".format(outlet).strip().lower()

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
                    user_input = option_limit(action_req, input(": "))

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
                    user_input = option_limit(action_req, input(": "))

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
        tabox_df = pd.read_excel(local_database_filename, sheet_name=takeaway_box_sheetname)
        drink_df = pd.read_excel(local_database_filename, sheet_name=drink_stock_sheetname)

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
        try:
            drink_sheet = google_auth.open_by_url(tabox_url)
            rcv_sheetname_index = tabox_drink_sheet.worksheet(property="title", value=receivers_sheetname).index
            rcv_df = tabox_drink_sheet[rcv_sheetname_index].get_as_df()
        except:
            try:
                rcv_url = tabox_url + "htmlview"
                rcv_df = pd.read_html(rcv_url, encoding="utf-8")[2]
                parseGoogleHTMLSheet(rcv_df)
            except:
                rcv_df = pd.read_excel(local_database_filename, sheet_name=receivers_sheetname)
    else:
        rcv_df = pd.read_excel(local_database_filename, sheet_name=receivers_sheetname)

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
            print(e)
            print("网络上传数据库可能失败！")
    else:
        print("无网络连接，数据将会暂时本地保存")
        print("待网络连接后，本地暂时保存的数据将会自动上传到云端")

    if wifi:
        try:
            online_db_sheet = google_auth.open_by_url(online_database_url)
            shift_db_sheet_index = online_db_sheet.worksheet(property="title", value=shift_db_sheetname).index
            shift_db = online_db_sheet[shift_db_sheet_index].get_as_df()

        except Exception as e:
            print()
            print("通过机器人从网络获取排班数据失败! 错误描述如下: ")
            print(e)
            print()
            shift_db = pd.read_excel(local_database_filename, sheet_name=shift_db_sheetname)
            print("已采用本地备份的排班数据")
    else:
        shift_db = pd.read_excel(local_database_filename, sheet_name=shift_db_sheetname)


    with pd.ExcelWriter(local_database_filename) as writer:
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

def parse_sending(on_duty, google_auth, outlet, send_dict, drink_message_string, tabox_message_string, print_result, k_dict, fernet_key, wifi, date_dict, value_dict, db_write):
    date = date_dict["dfb"].strftime("%Y-%m-%d")

    drink_stock_alert = eval(k_dict["drink_stock_alert"].strip().capitalize())
    box_stock_alert = eval(k_dict["box_stock_alert"].strip().capitalize())
    night_audit_alert = eval(k_dict["night_audit_alert"].strip().capitalize())

    send_drink_msg = send_dict["send_drink_msg"]
    send_tabox_msg = send_dict["send_tabox_msg"]
    send_print_result = send_dict["send_print_result"]

    night_audit_send_channel = k_dict["night_audit_send_channel"]
    drink_send_channel = k_dict["drink_send_channel"]
    tabox_send_channel = k_dict["tabox_send_channel"]

    rcv_sheetname = k_dict["receivers_sheetname"]
    box_drink_in_out_url = fernet_decrypt(k_dict["box_drink_in_out_url"], fernet_key)

    on_duty = str(on_duty).strip().upper()

    write_finance_db = db_write["write_finance_db"]

    try:
        box_drink_sheet = google_auth.open_by_url(box_drink_in_out_url)
        rcv_sheetname_index = box_drink_sheet.worksheet(property="title", value=rcv_sheetname).index
        rcv_df = box_drink_sheet[rcv_sheetname_index].get_as_df()
    except:
        try:
            rcv_url = box_drink_in_out_url + "htmlview"
            rcv_df = pd.read_html(rcv_url, encoding="utf-8")[2]
            parseGoogleHTMLSheet(rcv_df)
        except:
            local_database_filename = k_dict["local_database_filename"]
            rcv_df = pd.read_excel(local_database_filename, sheet_name=rcv_sheetname)


    rcv_dict = {}
    for index in range(len(rcv_df)):
        rcv_dict.update({ str(rcv_df.iloc[index, 0]) : str(rcv_df.iloc[index, 1]) })

    if drink_stock_alert:

        drink_stock_email_server = fernet_decrypt(k_dict["drink_stock_email_server"], fernet_key)
        drink_stock_email_sender = fernet_decrypt(k_dict["drink_stock_email_sender"], fernet_key)
        drink_stock_sender_password = fernet_decrypt(k_dict["drink_stock_sender_password"], fernet_key)
        drink_stock_telegram_bot_api = fernet_decrypt(k_dict["drink_stock_telegram_bot_api"], fernet_key)

        if send_drink_msg:
            if len(drink_message_string) > 0:
                if drink_send_channel.strip().capitalize() == "Telegram":
                    receivers = ast.literal_eval(rcv_dict["drink_telegram_receivers"])[on_duty]
                    sending_telegram(is_pr=False,
                                      message=drink_message_string,
                                      api = drink_stock_telegram_bot_api,
                                      receiver=receivers,
                                      wifi=wifi)
                elif drink_send_channel.strip().capitalize() == "Email":
                    receivers = fernet_decrypt(rcv_dict["drink_mail_receivers"], fernet_key).strip().split(",")
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
                    receivers = ast.literal_eval(rcv_dict["box_telegram_receivers"])[on_duty]
                    #receivers = ast_literal_eval(str(rcv_dict["box_telegram_receivers"]))[on_duty]
                    sending_telegram(is_pr=False,
                                      message=tabox_message_string,
                                      api = box_stock_telegram_bot_api,
                                      receiver=receivers,
                                      wifi=wifi)
                elif tabox_send_channel.strip().capitalize() == "Email":
                    receivers = fernet_decrypt(rcv_dict["box_mail_receivers"], fernet_key).strip().split(",")
                    #receivers = ast.literal_eval(str(rcv_dict["box_mail_receivers"]))[on_duty]
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
                    #receivers = fernet_decrypt(rcv_dict["night_audit_telegram_receivers"], fernet_key).strip().split(",")
                    receivers = ast.literal_eval(str(rcv_dict["night_audit_telegram_receivers"]))[on_duty]
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

                elif night_send_channel.strip().capitalize() == "Email":
                    print_result += ["——————————————", "服务费: ${}".format(svc),
                                     "GST: ${}".format(gst), "日均营业额: ${}".format(ads)]
                    #receivers = fernet_decrypt(rcv_dict["night_audit_mail_receivers"], fernet_key).strip().split(",")
                    receivers = ast.literal_eval(str(rcv_dict["night_audit_mail_receivers"]))[on_duty]
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
                        user_input = option_limit(action_req, input(": "))

                        if user_input == 0:
                            if night_audit_send_channel.strip().capitalize() == "Telegram":
                                #receivers = fernet_decrypt(rcv_dict["night_audit_telegram_receivers"], fernet_key).strip().split(",")
                                receivers = ast.literal_eval(str(rcv_dict["night_audit_telegram_receivers"]))[on_duty]
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

                            elif night_send_channel.strip().capitalize() == "Email":
                                print_result += ["——————————————", "服务费: ${}".format(svc),
                                                 "GST: ${}".format(gst), "日均营业额: ${}".format(ads)]
                                #receivers = fernet_decrypt(rcv_dict["night_audit_mail_receivers"], fernet_key).strip().split(",")
                                receivers = ast.literal_eval(str(rcv_dict["night_audit_mail_receivers"]))[on_duty]
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

def parse_display_df(value_dict, rule_df_dict, k_dict, google_auth, db_writables, fernet_key, take_databases, database_url):
    tabox_inv_function = eval(k_dict["takeaway_box_inventory"].strip().capitalize())
    drink_inv_function = eval(k_dict["drink_inventory"].strip().capitalize())

    bookkeeping_columns = ["描述", "收入", "支出", "月累计"]
    drink_columns = ["饮料(单位)", "入库", "出库", "剩余"]
    tabox_columns = ["物品(单位)", "入库", "用出", "剩余"]

    write_finance_db = db_writables["write_finance_db"]
    tabox_write_db = db_writables["tabox_write_db"]
    write_drink_db = db_writables["write_drink_db"]
    write_promo_db = db_writables["write_promo_db"]

    box_drink_in_out_url = fernet_decrypt(k_dict["box_drink_in_out_url"], fernet_key)
    tabox_sheetname = k_dict["takeaway_box_sheetname"]
    drink_stock_sheetname = k_dict["drink_stock_sheetname"]

    local_database_filename = k_dict["local_database_filename"]

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
                    tabox_df = pd.read_excel(local_database_filename, sheet_name=tabox_sheetname)


            try:
                tabox_db_sheet = google_auth.open_by_url(database_url)
                tabox_db_sheetname_index = tabox_db_sheet.worksheet(property="title", value=takeaway_box_database_sheetname).index
                tabox_db = tabox_db_sheet[tabox_db_sheetname_index].get_as_df()

            except:
                tabox_db = pd.read_excel(local_database_filename, sheet_name = takeaway_box_database_sheetname)

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
                    drink_df = pd.read_excel(local_database_filename, sheet_name=drink_stock_sheetname)

            try:
                drink_db_sheet = google_auth.open_by_url(database_url)
                drink_db_sheetname_index = drink_db_sheet.worksheet(property="title", value=drink_db_sheetname).index
                drink_db = drink_db_sheet[drink_db_sheetname_index].get_as_df()
            except:
                drink_db = pd.read_excel(local_database_filename, sheet_name=drink_db_sheetname)

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

def backup_script(script_backup_filename, script):
    with open("fernet_key.txt", "rb") as keyfile:
        fernet_key = keyfile.read()

    f_handler = Fernet(fernet_key)
    encrypted_script = f_handler.encrypt(script.encode())
    
    with open(script_backup_filename, "wb") as sfile:
        sfile.write(encrypted_script)    
    
def night_audit_main(on_duty, database_url, db_setting_url, serialized_rule_filename, service_filename, constants_sheetname, google_auth, box_num, drink_num, promo_num, lun_sales, lun_gc, tb_sales, tb_gc, lun_fwc, lun_kwc, tb_fwc, tb_kwc, night_fwc, night_kwc, script_backup_filename, script, wifi):
    on_duty = on_duty
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

    fernet_key = get_key()

    if fernet_key == 0:
        print("安全密钥错误! ")
    else:
        backup_script(script_backup_filename, script)

        if wifi:
            print("云端模式, 程序运行中, 请耐心等待...")
        else:
            print("离线模式, 程序运行中, 请耐心等待...")
        with tqdm(total=100) as pbar:
            outlet = get_outlet()
            k_dict, rule_df_dict = get_dfs(google_auth, db_setting_url, service_filename, constants_sheetname, serialized_rule_filename, outlet, fernet_key, wifi)
            book_dict = get_book_dfs(k_dict)

            if len(book_dict) == 0:
                pass
            else:
                date_dict = get_df_date(book_dict)
                pbar.update(3)

                get_columns = get_db_columns(k_dict, drink_num, box_num, promo_num)
                pbar.update(3)

                take_databases = retrieve_database(k_dict, google_auth, fernet_key, wifi, database_url)
                pbar.update(3)

                take_databases = update_columns(take_databases, get_columns, k_dict)
                pbar.update(3)

                tabox_write_db, box_value, box_unparsed_alert = parse_tabox(k_dict, google_auth, fernet_key, box_num, rule_df_dict, take_databases, book_dict, date_dict, wifi)
                pbar.update(10)

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
                                                 alert_type="box")

                pbar.update(10)
                value_dict, write_finance_db, write_promo_db = parse_value_dict(promo_num, lun_sales, lun_gc, tb_sales, tb_gc, lun_fwc, lun_kwc, tb_fwc, tb_kwc, night_fwc, night_kwc, rule_df_dict, take_databases, date_dict, book_dict, k_dict)
                pbar.update(4)

                write_drink_db, drink_dict, drink_unparsed_alert = parse_drink_stock(k_dict, fernet_key, google_auth, drink_num, value_dict, take_databases, date_dict, wifi)
                pbar.update(4)

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
                                                   alert_type="drink")

                pbar.update(12)

                print_result = parse_print_rule(value_dict, rule_df_dict, write_finance_db, write_promo_db, write_drink_db, tabox_write_db)
                pbar.update(5)

                db_writables = db_write(write_drink_db, tabox_write_db, write_finance_db, write_promo_db)
                pbar.update(5)

                take_databases, send_dict = pending_upload_db(take_databases, k_dict, box_value, drink_dict, value_dict, db_writables, date_dict, drink_num, promo_num, get_columns)
                pbar.update(5)

                upload_db(database_url, take_databases, k_dict, fernet_key, google_auth, box_num, drink_num, wifi, outlet)
                pbar.update(5)

                parse_sending(on_duty, google_auth, outlet, send_dict, drink_message_string, tabox_message_string, print_result, k_dict, fernet_key, wifi, date_dict, value_dict, db_writables)
                pbar.update(5)

                bk_df, show_box_df, show_drink_df, tabox_max_date, drink_db_max_date = parse_display_df(value_dict, rule_df_dict, k_dict, google_auth, db_writables, fernet_key, take_databases, database_url)
                pbar.update(5)

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

                if not show_box_df.empty:
                    print("酒水信息: ")
                    print(drink_db_max_date.strftime("%Y-%m-%d"))
                    prtdf(show_drink_df)
                    print()

                if not show_drink_df.empty:
                    print("打包盒信息: ")
                    print(tabox_max_date.strftime("%Y-%m-%d"))
                    prtdf(show_box_df)
                    print()


if __name__ == "__main__":
    night_audit_main(on_duty, database_url, db_setting_url, serialized_rule_filename, service_filename, constants_sheetname, google_auth, box_num, drink_num, promo_num, lun_sales, lun_gc, tb_sales, tb_gc, lun_fwc, lun_kwc, tb_fwc, tb_kwc, night_fwc, night_kwc, script_backup_filename, script, wifi)
