#! /usr/bin/env python3
#encoding='GBK'

import numpy as np
import pandas as pd
import string
import re
import math
import datetime as dt
import os
import unicodedata
import smtplib
from email.mime.text import MIMEText
from email.header import Header
import requests
import validators
from requests.structures import CaseInsensitiveDict
import io
#import urllib

def single_zero(x):
    if float(x) == 0:
        return int(0)
    else:
        return x

def resetAxis(df, axis):
    if axis == 0:
        df.reset_index(inplace=True)
        df.drop('index', axis=1, inplace=True)
        return df
    else:
        resetColumns = np.arange(0, len(df.columns)).astype(str)
        df.columns = resetColumns
        return df

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

        resetAxis(dfSlice, axis=1)

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

def promo_add_rule(databaseFileName, promoRecordSheetName, td_pd, td_ytd, promo_count, title, take):
    promo_df = pd.read_excel(databaseFileName, thousands=',', sheet_name=promoRecordSheetName)
    td_info_exists = promo_df[promo_df['Date'] == td_pd]

    weekday_dict = {
        '0' :'Sunday',
        '1' : 'Monday',
        '2' : 'Tuesday',
        '3' : 'Wednesday',
        '4' : 'Thursday',
        '5' : 'Friday',
        '6' : 'Saturday',
    }

    if td_info_exists.empty:
        if promo_df['Date'].values[-1] == td_ytd:
            if int(str(take)[0]) == 1:
                if dt.datetime.strptime(td_pd,'%Y-%m-%d').strftime("%A") == weekday_dict[str(take)[2]]:
                    promo_count_ytd = 0
                else:
                    promo_count_ytd = promo_df['{}_tdy'.format(title)].values[-1]

                if str(take)[3] == '1':
                    promo_count_tdy = promo_count_ytd + promo_count
                elif str(take)[3] == '2':
                    promot_count_tdy = promo_count_ytd - promo_count
            elif int(str(take)[0]) == 2:
                promo_count_ytd = promo_df['{}_tdy'.format(title)].values[-1]
                promo_count_tdy = promo_count_ytd + promo_count
            elif int(str(take)[0]) == 3:
                promo_count_ytd = promo_df['{}_tdy'.format(title)].values[-1]
                promo_count_tdy = promo_count_ytd - promo_count
        else:
            promo_count_ytd = np.nan
            promo_count_tdy = np.nan
    else:
        promo_count_ytd = int(td_info_exists['{}_ytd'.format(title)].values[0])
        promo_count_tdy = int(td_info_exists['{}_tdy'.format(title)].values[0])


    return [promo_count_ytd, promo_count_tdy]

def ModelUseSums(dataFrame, keyName, modelListing, TBRuleDf, FPara):
    df = dataFrame.copy()
    df = df[df['0'].isin(modelListing[str(keyName)])].dropna(axis=1)

    if df.empty:
        return float(0)
    else:
        df = df.iloc[:, :2]

        ruleDf = TBRuleDf.copy()
        ruleDf['Model to Use'] = ruleDf['Model to Use'].apply(lambda x: FPara[str(x)])
        ruleDf = ruleDf[ruleDf['Model to Use'] == str(keyName)]
        ruleDf.drop(['Chinese food name', 'Model to Use', 'active status'], axis=1, inplace=True)
        ruleDf.columns = ['0', 'Multiplier']

        mergedDf = pd.merge(left=df, right=ruleDf, how='inner')

        mergedDf['1'] = mergedDf['1'].astype(float)
        mergedDf['Multiplier'] = mergedDf['Multiplier'].astype(float)

        mergedDf['After Multiplication'] = mergedDf['1'] * mergedDf['Multiplier']

        return float(mergedDf['After Multiplication'].sum())

def modelLeft(df, df_string, date):
    mleft = df[df['Date'] == date]['{}现有数量'.format(df_string)].values[0]
    return mleft

def model_now(df, key_name, box_dict, td_ytd):
    if float(box_dict['{}_absolute'.format(key_name)]) < 0:
        xianYouShuLiang = float(df[df['Date'] == td_ytd]['{}现有数量'.format(key_name)].values[0])
        ruKu = float(abs(box_dict['{}_in'.format(key_name)]))
        chuKu = float(abs(box_dict['{}_out'.format(key_name)]))
        chuShou = float(abs(box_dict['{}_sale'.format(key_name)]))

        value = float(xianYouShuLiang + ruKu - chuKu - chuShou)

    elif float(box_dict['{}_absolute'.format(key_name)]) >= 0:

        value = float(box_dict['{}_absolute'.format(key_name)])

    return value

def stock_alert(box_value, model_df, noti_df, td_pd, alert_everyday=False):
    day_name = dt.datetime.strptime(td_pd,'%Y-%m-%d').strftime("%A")
    alert_dict = {}
    if alert_everyday:
        for index in range(len(model_df)):
            alert_dict.update({
                    '{}'.format(model_df.iloc[index, 7]) : box_value['{}_now'.format(model_df.iloc[index, 0])]
            })

    elif not alert_everyday:
        if day_name == 'Friday':
            for index in range(len(model_df)):
                alert_dict.update({
                        '{}'.format(model_df.iloc[index, 7]) : box_value['{}_now'.format(model_df.iloc[index, 0])]
                })

        elif day_name != 'Friday':
            for index in range(len(model_df)):
                if box_value['{}_now'.format(model_df.iloc[index, 0])] is np.nan:
                    continue
                else:
                    if float(box_value['{}_now'.format(model_df.iloc[index, 0])]) <= float(model_df.iloc[index, 6]):
                        alert_dict.update({
                            '{}'.format(model_df.iloc[index, 0]) : box_value['{}_now'.format(model_df.iloc[index, 0])]
                        })
                    else:
                        continue

    noti_alert = {}
    for index in range(len(noti_df)):
        if str(noti_df.iloc[index, 1]) == td_pd:
            noti_alert.update({
                    '{}'.format(noti_df.iloc[index, 0]) : '叫货提醒日期{}'.format(noti_df.iloc[index, 1])
            })

        else:
            continue

    return [alert_dict, noti_alert]

def sending_email(text, receiver_list, td_pd, emailSender, emailSenderPassword):
    mail_server = 'smtp.126.com'

    message = MIMEText(str(text), 'plain', 'utf-8')
    message['From'] = emailSender

    subject = '物品库存通知 {}'.format(td_pd)
    message['Subject'] = Header(subject, 'utf-8')

    try:
        smtpObj = smtplib.SMTP()
        smtpObj.connect(mail_server, 25)
        smtpObj.login(emailSender, emailSenderPassword)
        for ppl in receiver_list:
            print('正在发送电邮给{}'.format(ppl))
            message['To'] = ppl
            smtpObj.sendmail(emailSender, ppl, message.as_string())
            print('{} 电邮发送成功'.format(ppl))
    except smtplib.SMTPException:
        print('电邮发送失败')

def normal_round(n, decimals=0):
    expoN = n * 10 ** decimals
    if abs(expoN) - abs(math.floor(expoN)) < 0.5:
        return math.floor(expoN) / 10 ** decimals
    return math.ceil(expoN) / 10 ** decimals

def remove_spaces(var):
    return str(var).translate({ord(y): None for y in string.whitespace})

#def on_internet():
#    url = "https://www.google.com.sg/"
#    timeout = 5
#    try:
#        request = requests.get(url, timeout=timeout)
#        return True
#    except (requests.ConnectionError, requests.Timeout) as exception:
#        return False

def drink_beizhu(ziyong):
    if ziyong == 0:
        return None
    else:
        return str(ziyong) + '罐/瓶自用'

def prtdf(df):
    with pd.option_context('display.max_rows', None,
                           'display.max_columns', None,
                           'display.width', 1000,
                           'display.precision', 2,
                           'display.colheader_justify', 'center'):
        return display(df)

def readCsv(githubUserName, githubRepoName, githubBranchName, githubFileName, csvSep, csvEncoding, runLocally):
    if runLocally:
        return pd.read_csv(githubFileName, sep=csvSep, encoding=csvEncoding)
    else:
        githubPrefix = "https://raw.githubusercontent.com"
        URL = "{}/{}/{}/{}/{}".format(githubPrefix, githubUserName, githubRepoName, githubBranchName, githubFileName)

        resp = requests.get(URL)
        if resp.status_code == 200:
            return pd.read_csv(io.StringIO(resp.text), sep=csvSep, encoding=csvEncoding)
        else:
            raise ConnectionError("访问错误")

#中午营业额
#lun_sales = 0

#中午顾客人数
#lun_gc = 0

#中午楼面员工人数
#lun_fwc = '4'

#中午厨房员工人数
#lun_kwc = '7+1洗碗工'

#下午茶营业额
#tb_sales = 0

#下午茶顾客
#tb_gc = 0

#下午茶楼面员工人数
#tb_fwc = '1'

#下午茶厨房员工人数
#tb_kwc = '0'

#晚上楼面员工人数
#night_fwc = '4'

#晚上厨房员工人数
#night_kwc = '7+1洗碗工'


#--------酒水分割线------------#

#可乐入库数量
#coke_in = 0
#可乐自用数量
#coke_ziyong = 0

#老虎入库数量
#tiger_in = 0
#老虎自用数量
#tiger_ziyong = 0

#喜力入库数量
#heineken_in = 0
#喜力自用数量
#heineken_ziyong = 0

#矿泉水入库数量
#water_in = 0
#矿泉水自用数量
#water_ziyong = 0

#红酒入库数量
#redwine_in = 0
#红酒自用数量
#redwine_ziyong = 0

#白酒入库数量
#whitewine_in = 0
#白酒自用数量
#whitewine_ziyong = 0

#圣诞套餐红酒入库数量
#xmas_wine_in = 0
#圣诞套餐红酒自用数量
#xmas_wine_ziyong = 0

#--------现金信息分割线----------#

#未存款现金的起始日期(格式: 年年年年-月月-日日)
#现金袋里目前有的最早日期
#cash_start_date = '2022-11-01'

#未存款现金的结束日期(格式: 年年年年-月月-日日, 默认值: 'default')
#'default'是根据起始日期直到今天为止计算现金收入总金额
#结束日期不能早于起始日期
#cash_end_date = 'default'

#githubUserName = ""
#githubRepoName = ""
#githubBranchName = ""
#emailSender = ''
#emailSenderPassword = ''
#emailReceiverList = ['']
#takeawayBoxInventoryURL = ""

functionalityRuleCsvName = "FunctionalityRule.csv"
valueDictRuleCsvName = 'ValueDictRule.csv'
printRuleCsvName = 'PrintRule.csv'
takeawayBoxRuleCsvName = 'TakeawayBoxRule.csv'
extraTakeawayBoxRuleCsvName = 'extraTakeawayBoxRule.csv'
NameConverterForShowingDfCsvName = 'NameConverterForShowingDf.csv'
promotionCountConstant = 30

drinkTakeOverName0In = int(abs(coke_in))
drinkTakeOverName1In = int(abs(tiger_in))
drinkTakeOverName2In = int(abs(heineken_in))
drinkTakeOverName3In = int(abs(water_in))
drinkTakeOverName4In = int(abs(redwine_in))
drinkTakeOverName5In = int(abs(whitewine_in))
drinkTakeOverName6In = int(abs(xmas_wine_in))
drinkTakeOverName7In = 0
drinkTakeOverName8In = 0
drinkTakeOverName9In = 0
drinkTakeOverName10In = 0
drinkTakeOverName11In = 0

drinkTakeOverName0Ziyong = int(abs(coke_ziyong))
drinkTakeOverName1Ziyong = int(abs(tiger_ziyong))
drinkTakeOverName2Ziyong = int(abs(heineken_ziyong))
drinkTakeOverName3Ziyong = int(abs(water_ziyong))
drinkTakeOverName4Ziyong = int(abs(redwine_ziyong))
drinkTakeOverName5Ziyong = int(abs(whitewine_ziyong))
drinkTakeOverName6Ziyong = int(abs(xmas_wine_ziyong))
drinkTakeOverName7Ziyong = 0
drinkTakeOverName8Ziyong = 0
drinkTakeOverName9Ziyong = 0
drinkTakeOverName10Ziyong = 0
drinkTakeOverName11Ziyong = 0

#print("测试网络环境...")
#on_net = on_internet()

#if not on_net:
#    print('没有网络连接')
#else:
#    print("网路连接正常")

#URL = "{}/{}/{}/{}/{}".format("https://raw.githubusercontent.com", githubUserName, githubRepoName, githubBranchName, githubFileName)
#script = urllib.request.urlopen(URL).read().decode()
#exec(script)

functionRuleDf = readCsv(githubFileName=functionalityRuleCsvName,
                         githubUserName=githubUserName,
                         githubRepoName=githubRepoName,
                         githubBranchName=githubBranchName,
                         csvSep='|',
                         csvEncoding='utf-8',
                         runLocally=FPara['runLocally'])

FPara = {}
for row in range(len(functionRuleDf)):
    if str(functionRuleDf.iloc[row, 1]) in ['True', 'False']:
        FPara.update({ str(functionRuleDf.iloc[row, 0]) : eval(str(functionRuleDf.iloc[row, 1])) })
    else:
        FPara.update({ str(functionRuleDf.iloc[row, 0]) : str(functionRuleDf.iloc[row, 1]) })

FPara['emailReceiverList'] = emailReceiverList
FPara['emailSender'] = emailSender
FPara['emailSenderPassword'] = emailSenderPassword
FPara['githubUserName'] = githubUserName
FPara['githubRepoName'] = githubRepoName
FPara['githubBranchName'] = githubBranchName
FPara['takeawayBoxInventoryURL'] = takeawayBoxInventoryURL
FPara['runLocally'] = runLocally

takeawayBoxInventoryFunction = FPara['takeawayBoxInventoryFunction']
stockAlertFunction = FPara['stockAlertFunction']
stockEverydayFunction = FPara['stockEverydayFunction']
cashRecordShowFunction = FPara['cashRecordShowFunction']
drinkInventoryFunction = FPara['drinkInventoryFunction']


chineseFunction = ["打包盒库存功能", "打包盒库存电邮提醒功能",
                   "每日打包盒库存电邮提醒功能",
                   "显示每日现金记录功能", "酒水库存功能",]
functionVarList = [takeawayBoxInventoryFunction,
                   stockAlertFunction,
                   stockEverydayFunction,
                   cashRecordShowFunction,
                   drinkInventoryFunction]

print()
for index in range(len(chineseFunction)):
    if functionVarList[index]:
        print("{}已开启".format(chineseFunction[index]))
    else:
        print("{}已关闭".format(chineseFunction[index]))

print()
drinkCounter = 0
boxCounter = 0
for key in FPara.keys():
    if key.startswith("drinkTakeOverName"):
        drinkCounter += 1
    elif key.startswith("boxTakeOverName"):
        boxCounter += 1
    else:
        continue

databaseColumns = {}

drinkRecordColumns = ['日期']
for num in range(drinkCounter):
    drinkRecordColumns += [
                           '{}上日存货'.format(FPara['drinkTakeOverName{}'.format(num)]),
                           '{}进'.format(FPara['drinkTakeOverName{}'.format(num)]),
                           '{}出'.format(FPara['drinkTakeOverName{}'.format(num)]),
                           '{}实结存'.format(FPara['drinkTakeOverName{}'.format(num)]),
                           '{}备注'.format(FPara['drinkTakeOverName{}'.format(num)])
                          ]


databaseColumns[FPara['drinkRecordSheetName']] = drinkRecordColumns

boxRecordColumns = ['Date']
for num in range(boxCounter):
    boxRecordColumns += [
                         '{}入库'.format(FPara['boxTakeOverName{}'.format(num)]),
                         '{}出库'.format(FPara['boxTakeOverName{}'.format(num)]),
                         '{}现有数量'.format(FPara['boxTakeOverName{}'.format(num)])

                        ]

databaseColumns[FPara['takeawayBoxRecordSheetName']] = boxRecordColumns

read = pd.read_excel(FPara['bookFileName'],thousands=',')
columnsRename = np.arange(0, len(read.columns)).astype(str)
read.columns = columnsRename
read['0'] = read['0'].astype(str)

if len(read) > int(FPara['minBookFileLenAllowable']):
    try:
        bi = int(read[read['0'].str.contains('Date:')].index.values[0])
        ei = int(read[read['0'] == 'Total Sales'].index.values[0])
        outlet_loc = str(read.iloc[(bi+1):(ei),:].dropna(axis=1).values[0][0])
        continue_ = True
    except IndexError:
        continue_ = False

    if continue_:
        pbi = int(read[read['0'].str.contains('PAYMENT BREAKDOWN:')].index[0])

        try:
            pei = int(read[read['0'].str.contains('PAYMENT BREAKDOWN (POS):', regex=False)].index[0])
        except IndexError:
            pei = int(read[read['0'].str.contains('Transaction Void Items')].index[0])

        pay_breakdown = read.copy()
        pay_breakdown = pay_breakdown.iloc[(pbi+1):(pei), :]
        resetAxis(pay_breakdown, axis=0)
        pay_breakdown['0'] = pay_breakdown['0'].astype(str)

        dropColumns = []
        for columnIndex in range(len(pay_breakdown.columns)):
            if pay_breakdown[str(columnIndex)].isnull().all():
                dropColumns += [str(columnIndex)]
            else:
                continue

        pay_breakdown.drop(dropColumns, axis=1, inplace=True)
        resetAxis(pay_breakdown, axis=0)
        resetAxis(pay_breakdown, axis=1)
        pay_breakdown = pay_breakdown[pay_breakdown['0'] != 'nan']

        if pay_breakdown.empty:
            print("无任何付款信息")

        pay_brkdwn = pay_breakdown.copy()
        pay_brkdwn['0'] = pay_brkdwn['0'].apply(lambda x:x[:x.find('(')])

        read_ = read.copy()
        read_['0'] = read_['0'].apply(lambda a : unicodedata.normalize("NFKD", a))
        read_['0'] = read_['0'].apply(lambda x : remove_spaces(x))

        df_dict = {}
        df_dict.update({
            'read_' : read_,
            'read' : read,
            'pay_breakdown': pay_breakdown,
            'pay_brkdwn' : pay_brkdwn,
            })

        today_dt_format = dt.datetime.today()
        today_date = today_dt_format.strftime("%Y年%m月%d日")

        #raw_date_from_book
        rdfb = read[read['0'].str.contains('Date:')].dropna(axis=1).values[0][0]
        rdfb = rdfb.split()
        if len(rdfb) == 8:
            if rdfb[1] == rdfb[5] and rdfb[2] == rdfb[6] and rdfb[3] == rdfb[7]:
                #date_from_book
                dfb = '{} {} {}'.format(rdfb[1],rdfb[2],rdfb[3])
                dfb = dt.datetime.strptime(dfb, '%d %b %Y')
                take_date = dfb.strftime("%Y年%m月%d日")
            else:
                print('不支持跨日期。\n报表日期将以今天日期为标准。')
                dfb = dt.datetime.today().strptime(dfb, '%d %b %Y')
                take_date = today_date
        else:
            print('日期格式不符，报表将以今天的日期为标准。')
            dfb = dt.datetime.today().strptime(dfb, '%d %b %Y')
            take_date = today_date

        td_pd = dfb.strftime('%Y-%m-%d')
        td_ytd = (dfb - dt.timedelta(days=1)).strftime('%Y-%m-%d')
        yesterday_date = (dfb - dt.timedelta(days=1)).strftime('%Y年%m月%d日')

        TBRuleDf = readCsv(githubFileName=takeawayBoxRuleCsvName,
                       githubUserName=FPara['githubUserName'],
                       githubRepoName=FPara['githubRepoName'],
                       githubBranchName=FPara['githubBranchName'],
                       csvSep='|',
                       csvEncoding='utf-8',
                       runLocally=FPara['runLocally'])

        TBRuleDf['String Locator'] = TBRuleDf['String Locator'].apply(lambda a : unicodedata.normalize("NFKD", a))
        TBRuleDf['String Locator'] = TBRuleDf['String Locator'].apply(lambda x : remove_spaces(x))
        TBRuleDf['active status'] = TBRuleDf['active status'].astype(int)

        model_listing = {}
        for number in range(boxCounter):
            modelName = FPara['boxTakeOverName{}'.format(number)]
            model_listing.update({ modelName : TBRuleDf[(TBRuleDf['Model to Use'] == 'boxTakeOverName{}'.format(number)) & (TBRuleDf['active status'] == 1)]['String Locator'].values.tolist()})

        html_df = pd.read_html(FPara['takeawayBoxInventoryURL'], encoding='utf-8')
        model_df = html_df.copy()[0]
        model_df.drop('Unnamed: 0', axis=1, inplace=True)
        model_df.columns = model_df.iloc[0, :]
        model_df.drop(0, axis=0, inplace=True)
        model_df.reset_index(inplace=True)
        model_df.drop('index', axis=1, inplace=True)


        noti_df = html_df.copy()[1]
        noti_df.drop('Unnamed: 0', axis=1, inplace=True)
        noti_df.columns = noti_df.iloc[0,:]
        noti_df.drop(0, axis=0, inplace=True)
        noti_df.reset_index(inplace=True)
        noti_df.drop('index', axis=1, inplace=True)

        for i in noti_df.columns:
            noti_df[i] = noti_df[i].astype(str)

        noti_df[noti_df.columns[-1]] = noti_df[noti_df.columns[-1]].apply(lambda x : str(x[x.find('(')+1:x.find(')')]) )

        tabox = pd.read_excel(FPara['databaseFileName'], sheet_name=FPara['takeawayBoxRecordSheetName'])

        if len(tabox.columns) == len(databaseColumns[FPara['takeawayBoxRecordSheetName']]):
            if tabox.columns.tolist() != databaseColumns[FPara['takeawayBoxRecordSheetName']]:
                tabox.columns = databaseColumns[FPara['takeawayBoxRecordSheetName']]
                print("The columns of {} had been updated.".format(FPara['takeawayBoxRecordSheetName']))

        tabox_tdy_record = tabox[tabox['Date'] == td_pd]

        if takeawayBoxInventoryFunction:
            box_value = {}
            for index in range(len(model_df)):
                box_value.update({
                    '{}_in'.format(model_df.iloc[index, 0]) : int(model_df.iloc[index, 1])*float(eval(str(model_df.iloc[index, 5]))),

                    '{}_out'.format(model_df.iloc[index, 0]) : int(model_df.iloc[index, 2])*float(eval(str(model_df.iloc[index, 5]))),

                    '{}_absolute'.format(model_df.iloc[index, 0]) : float(model_df.iloc[index, 3]),

                    '{}_sale'.format(model_df.iloc[index, 0]): float(normal_round(float(ModelUseSums(dataFrame=df_dict['read_'],
                                                                                                keyName=str(model_df.iloc[index, 0]),
                                                                                                modelListing=model_listing,
                                                                                                TBRuleDf = TBRuleDf, FPara = FPara))*
                                                               float(model_df.iloc[index, 4])*
                                                               float(eval(str(model_df.iloc[index, 5]))), 4))
                })

            TBExtraRuleDf = readCsv(githubFileName=extraTakeawayBoxRuleCsvName,
                           githubUserName=FPara['githubUserName'],
                           githubRepoName=FPara['githubRepoName'],
                           githubBranchName=FPara['githubBranchName'],
                           csvSep='|',
                           csvEncoding='utf-8',
                           runLocally=FPara['runLocally'])

            if TBExtraRuleDf.empty:
                pass

            else:
                TBExtraRuleDf['active status'] = TBExtraRuleDf['active status'].astype(int)
                for op in range(len(TBExtraRuleDf)):
                    if int(TBExtraRuleDf['active status'][op]) != 1:
                        continue
                    else:
                        if str(TBExtraRuleDf['operation'][op]) == 'plus equal':
                            box_value[str(TBExtraRuleDf['box_value to operate on'][op])] += box_value[str(TBExtraRuleDf['value'][op])]
                        elif str(TBExtraRuleDf['operation'][op]) == 'minus equal':
                            box_value[str(TBExtraRuleDf['box_value to operate on'][op])] -= box_value[str(TBExtraRuleDf['value'][op])]
                        elif str(TBExtraRuleDf['operation'][op]) == 'divide equal':
                            box_value[str(TBExtraRuleDf['box_value to operate on'][op])] /= box_value[str(TBExtraRuleDf['value'][op])]
                        elif str(TBExtraRuleDf['operation'][op]) == 'times equal':
                            box_value[str(TBExtraRuleDf['box_value to operate on'][op])] *= box_value[str(TBExtraRuleDf['value'][op])]
                        elif str(TBExtraRuleDf['operation'][op]) == 'equal':
                            box_value[str(TBExtraRuleDf['box_value to operate on'][op])] = box_value[str(TBExtraRuleDf['value'][op])]
                        elif str(TBExtraRuleDf['operation'][op]) == 'zero':
                            box_value[str(TBExtraRuleDf['box_value to operate on'][op])] = 0
                        elif str(TBExtraRuleDf['operation'][op]) == 'change value':
                            box_value[str(TBExtraRuleDf['box_value to operate on'][op])] = float(str(TBExtraRuleDf['value'][op]))
                        else:
                            pass

        elif not takeawayBoxInventoryFunction:
            box_value = {}
            for index in range(len(model_df)):
                box_value.update({
                    '{}_in'.format(model_df.iloc[index, 0]) : 0,
                    '{}_out'.format(model_df.iloc[index, 0]) : 0,
                    '{}_absolute'.format(model_df.iloc[index, 0]) : 0,
                    '{}_sale'.format(model_df.iloc[index, 0]): 0,
                })


        if tabox_tdy_record.empty:
            if tabox[tabox['Date'] == td_ytd].empty:
                tabox_write = False
                for index in range(len(model_df)):
                    box_value.update({
                        '{}_now'.format(model_df.iloc[index, 0]) : np.nan
                    })
            else:
                tabox_write = True

        else:
            tabox_write = False
            if takeawayBoxInventoryFunction:
                for index in range(len(model_df)):
                    box_value.update(
                        {
                            '{}_now'.format(model_df.iloc[index, 0]) : float(modelLeft(df=tabox,
                                                                                df_string=str(model_df.iloc[index, 0]),
                                                                                       date=td_pd))
                        })
            elif not takeawayBoxInventoryFunction:
                for index in range(len(model_df)):
                    box_value.update({
                            '{}_now'.format(model_df.iloc[index, 0]) : 0
                    }
                    )

        if tabox_write:
            if takeawayBoxInventoryFunction:
                for index in range(len(model_df)):
                    box_value.update({
                        '{}_now'.format(model_df.iloc[index, 0]) : float(model_now(df=tabox,
                                                                                   key_name=str(model_df.iloc[index, 0]),
                                                                                   box_dict=box_value,
                                                                                   td_ytd=td_ytd))
                    })
            elif not takeawayBoxInventoryFunction:
                for index in range(len(model_df)):
                    box_value.update({
                        '{}_now'.format(model_df.iloc[index, 0]) : 0
                    })

            update_tabox = {}
            tabox_input = [td_pd,]
            for x in range(len(model_df['title'])):
                tabox_input += [
                                box_value['{}_in'.format(model_df['title'][x])],
                                abs(box_value['{}_out'.format(model_df['title'][x])])+abs(box_value['{}_sale'.format(model_df['title'][x])]),
                                box_value['{}_now'.format(model_df['title'][x])]
                               ]

            for i in range(len(tabox_input)):
                update_tabox.update({tabox.columns[i]:[tabox_input[i]]})

            update_tabox = pd.DataFrame(update_tabox)
            tabox_concat = pd.concat([tabox, update_tabox])
            tabox_concat.sort_values(by='Date', ascending=True, ignore_index=True, inplace=True)
            tabox_msg = "{}的打包盒信息已存入数据库，写入时间: {}".format(take_date, str(pd.to_datetime('today')))

            if stockAlertFunction:
                if takeawayBoxInventoryFunction:
                    alerts = stock_alert(box_value=box_value, model_df=model_df, noti_df=noti_df, td_pd=td_pd, alert_everyday=stockEverydayFunction)
                    if len(alerts[0]) == 0 and len(alerts[1]) == 0:
                        send_email = False
                    elif len(alerts[0]) != 0 and len(alerts[1]) == 0:
                        send_email = True
                        text = '打包盒: \n'
                        for key, items in alerts[0].items():
                            text += '{}: 剩{} (条/个) \n'.format(key, round(float(items), 3))
                            text += '\n'

                    elif len(alerts[0]) == 0 and len(alerts[1]) != 0:
                        send_email = True
                        text = '{}提醒叫货的项目: \n'.format(td_pd)
                        for key, items in alerts[1].items():
                            text += '{}'.format(key)
                            text += '\n'

                    elif len(alerts[0]) != 0 and len(alerts[1]) != 0:
                        send_email=True

                        text = '打包盒: \n'
                        for key, items in alerts[0].items():
                            text += '{}: 剩{} (条/个) \n'.format(key, round(float(items), 3))
                            text += '\n'

                        text += '\n'

                        text += '{}提醒叫货的项目: \n'.format(td_pd)
                        for key, items in alerts[1].items():
                            text += '{}'.format(key)
                            text += '\n'

                    if send_email:
                        sending_email(text=text, receiver_list=FPara['emailReceiverList'], td_pd=td_pd, emailSender=FPara['emailSender'], emailSenderPassword=FPara['emailSenderPassword'])

                    elif not send_email:
                        pass

                elif not takeawayBoxInventoryFunction:
                    pass

            elif not stockAlertFunction:
                pass

        else:
            tabox_concat = tabox
            tabox_concat.sort_values(by='Date', ascending=True, ignore_index=True, inplace=True)
            tabox_msg = "{}的打包盒信息无法存入或无法再次存入数据库".format(take_date)

        print("计算中...请耐心等待")
        lun_sales = float(lun_sales)
        tb_sales = float(tb_sales)
        if lun_gc == 0:
            lun_avg = 0
        else:
            lun_avg = lun_sales/lun_gc

        if tb_gc == 0:
            tb_avg = 0
        else:
            tb_avg = tb_sales/tb_gc

        total_sales = float(resetAxis(read[read['0'] == 'Total Sales'].dropna(axis=1), axis=1)['1'].values[0])
        no_of_cover = int(resetAxis(read[read['0'] == 'No. Of Cover'].dropna(axis=1), axis=1)['1'].values[0])

        night_sales = total_sales - lun_sales - tb_sales
        night_gc = no_of_cover - lun_gc - tb_gc

        if night_gc == 0:
            night_avg = 0
        else:
            night_avg = night_sales/night_gc

        lun_avg = normal_round(lun_avg, 2)
        tb_avg = normal_round(tb_avg, 2)
        night_sales = normal_round(night_sales, 2)
        night_avg = normal_round(night_avg, 2)

        lun_sales = single_zero(format(lun_sales, '.2f'))
        lun_avg = single_zero(format(lun_avg, '.2f'))
        tb_sales = single_zero(format(tb_sales, '.2f'))
        tb_avg = single_zero(format(tb_avg, '.2f'))
        night_sales = single_zero(format(night_sales, '.2f'))
        night_avg = single_zero(format(night_avg, '.2f'))

        value_dict = {
            'take_date' : take_date,
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


        rule = readCsv(githubFileName=valueDictRuleCsvName,
                       githubUserName=FPara['githubUserName'],
                       githubRepoName=FPara['githubRepoName'],
                       githubBranchName=FPara['githubBranchName'],
                       csvSep='|',
                       csvEncoding='utf-8',
                       runLocally=FPara['runLocally'])

        for i in range(len(rule.index)):
            if int(rule.iloc[i, 7]) == 1:
                value_dict.update({
                    str(rule.iloc[i,0]): sum_rule(dfSlice=look_up_rule(dataFrame=df_dict[str(rule.iloc[i,2])],
                                                                       a_string=str(rule.iloc[i, 1]),
                                                                       take=int(rule.iloc[i, 3])),
                                                  take=int(rule.iloc[i, 4]),
                                                  is_money=int(rule.iloc[i,6]))
                })

                if int(rule.iloc[i, 5]) < 0:
                    continue
                else:
                    value_dict.update({
                        '{}_ytd_tdy'.format(str(rule.iloc[i,0]).replace('_count', '')) : promo_add_rule(
                                                                                     databaseFileName=FPara['databaseFileName'],
                                                                                     promoRecordSheetName=FPara['promoRecordSheetName'],
                                                                                     td_pd = td_pd,
                                                                                     td_ytd = td_ytd,
                                                                                     promo_count=int(value_dict[str(rule.iloc[i,0])]),
                                                                                     title=str(rule.iloc[i,0]).replace('_count', ''),
                                                                                     take=rule.iloc[i,5])
                    })

            else:
                continue

        value_dict.update({'net_sales_deductibles' : float(float(value_dict['TRUEBLUE_SUM'])+float(value_dict['SANRENXING_SUM'])+float(value_dict['PANDABOX_SUM'])+float(value_dict['Chinatown_5']))})
        value_dict.update({'net_sales_aft_deduction' : float(float(value_dict['net_sales']) - float(value_dict['net_sales_deductibles']))})
        value_dict['net_sales_aft_deduction'] = single_zero(normal_round(value_dict['net_sales_aft_deduction'], 2))

        read_net_sales = pd.read_excel(FPara['databaseFileName'], thousands=',', sheet_name=FPara['salesRecordSheetName'])
        dfb_record_exist = read_net_sales[read_net_sales['Date'] == td_pd]
        if dfb_record_exist.empty:
            if read_net_sales['Date'].values[-1] == td_ytd:
                if int(dfb.day) == 1:
                    cmns_ytd = 0.0
                    cmns_tdy = cmns_ytd + float(value_dict['net_sales_aft_deduction'])
                    cmns_tdy_ = normal_round(cmns_tdy, 2)
                    avg_daily_sales = cmns_tdy/int(dfb.day)
                    ads_rd1 = str(avg_daily_sales)[str(avg_daily_sales).find('.'):][:3]
                    ads_rd2 = str(avg_daily_sales)[:str(avg_daily_sales).find('.')]
                    ads = float(ads_rd2+ads_rd1)
                    ns_new_input = True
                else:
                    cmns_ytd = float(read_net_sales['cum net sales today'].values[-1])
                    cmns_tdy = cmns_ytd + float(value_dict['net_sales_aft_deduction'])
                    cmns_tdy_ = normal_round(cmns_tdy, 2)
                    avg_daily_sales = cmns_tdy/int(dfb.day)
                    ads_rd1 = str(avg_daily_sales)[str(avg_daily_sales).find('.'):][:3]
                    ads_rd2 = str(avg_daily_sales)[:str(avg_daily_sales).find('.')]
                    ads = float(ads_rd2+ads_rd1)
                    ns_new_input = True
            else:
                print('{}的昨日({})营业额信息不存在!'.format(td_pd, td_ytd))
                cmns_ytd = np.nan
                cmns_tdy_ = np.nan
                ads = np.nan
                ns_new_input = False
        else:
            print('{}的营业额信息已存在数据库'.format(td_pd))
            cmns_ytd = float(dfb_record_exist['cum net sales ytd'].values[0])
            cmns_tdy = float(dfb_record_exist['cum net sales today'].values[0])
            ads = float(dfb_record_exist['Avg Net Sales/Day'].values[0])
            cmns_tdy_ = format(cmns_tdy, '.2f')
            ns_new_input = False

        value_dict.update({
            'cmns_ytd' : cmns_ytd,
            'cmns_tdy_' : cmns_tdy_,
            'ads' : ads,
        })

        for number in range(promotionCountConstant):
            value_dict.update({'promo{}_ytd'.format(number+1) : value_dict['promo{}_ytd_tdy'.format(number+1)][0],
                               'promo{}_tdy'.format(number+1) : value_dict['promo{}_ytd_tdy'.format(number+1)][1]})

        promo_df = pd.read_excel(FPara['databaseFileName'], thousands=',', sheet_name=FPara['promoRecordSheetName'])
        promo_info_exist = promo_df[promo_df['Date'] == td_pd]
        if promo_info_exist.empty:
            if promo_df['Date'].values[-1] == td_ytd:
                sec_new = True
            else:
                print('{}的昨日({})促销信息不存在!'.format(td_pd, td_ytd))
                sec_new = False
        else:
            print('{}的促销信息已存在数据库'.format(td_pd))
            sec_new = False

        read_cash_record = pd.read_excel(FPara['databaseFileName'], thousands=",", sheet_name=FPara['cashRecordSheetName'])
        rcr_record_exist = read_cash_record[read_cash_record["Date"] == td_pd]
        if rcr_record_exist.empty:
            if read_cash_record["Date"].values[-1] == td_ytd:
                cash_received_tdy = float(value_dict["CASH"])
                cash_record_write = True
            else:
                print('{}的昨日({})现金记录不存在!'.format(td_pd, td_ytd))
                cash_received_tdy = None
                cash_record_write = False
        else:
            print("{}的现金记录已存在数据库".format(td_pd))
            cash_received_tdy = float(read_cash_record[read_cash_record["Date"] == td_pd]["Amount"].values[0])
            cash_record_write = False

        generate=True
        for num in range(promotionCountConstant):
            if int(value_dict['promo{}_ytd_tdy'.format(num+1)][1]) < int(value_dict['promo{}_count'.format(num+1)]):
                generate = False
            else:
                pass

        if generate:
            print_result = []
            printRuleDf = readCsv(githubFileName=printRuleCsvName,
                           githubUserName=FPara['githubUserName'],
                           githubRepoName=FPara['githubRepoName'],
                           githubBranchName=FPara['githubBranchName'],
                           csvSep='|',
                           csvEncoding='utf-8',
                           runLocally=FPara['runLocally'])

            for column in printRuleDf.columns:
                printRuleDf[column] = printRuleDf[column].astype(str)

            for index in range(len(printRuleDf)):
                formatList = printRuleDf['key names in value_dict'][index]
                formatListList = formatList.split(',')
                formatEval = []
                for formats in formatListList:
                    if formats == 'nan':
                        continue
                    else:
                        formatEval += [value_dict[str(formats)]]

                if printRuleDf['print_statement'][index] == 'nan':
                    print_result += [' ']

                elif int(printRuleDf['active_status'][index]) != 1:
                    pass
                else:
                    print_result += [printRuleDf['print_statement'][index].format(*formatEval)]

            print()
            print('生成报表中...')
            print('——————————————')
            print()

            for items in print_result:
                print(items)

            print()
            print('——————————————')
            print()
            print('服务费: ${}'.format(value_dict['svc']))
            print('GST: ${}'.format(value_dict['gst']))
            print('日均营业额: ${}'.format(value_dict['ads']))
            print()
            print("帐务簿记")
            print(take_date)
            bookkeeping_particulars = ['初始额', '销售额']

            voucher_expense_list = ['Trueblue会员', '三人行礼券', 'PandaBox', "唐人街首单减$5"]
            voucher_expense_dict_locator = ['TRUEBLUE_SUM', 'SANRENXING_SUM', 'PANDABOX_SUM', "Chinatown_5"]

            income_list = [' ',
                           format(float(value_dict['net_sales']), '.2f'), ]

            expense_list = [' ',
                            ' ', ]

            balance_list = [format(float(value_dict['cmns_ytd']), '.2f'),
                            format(normal_round(float(value_dict['cmns_ytd'])+float(value_dict['net_sales']),2), '.2f'), ]

            account_balance = float(balance_list[-1])
            for index in range(len(voucher_expense_dict_locator)):
                if value_dict[voucher_expense_dict_locator[index]] != 0:
                    bookkeeping_particulars += [voucher_expense_list[index]]
                    income_list += [' ']
                    expense_list += [format(float(value_dict[voucher_expense_dict_locator[index]]), '.2f')]
                    account_balance -= float(value_dict[voucher_expense_dict_locator[index]])
                    balance_list += [format(account_balance, '.2f')]
                else:
                    continue

            income_total = 0
            expense_total = 0
            for index in range(len(income_list)):
                if income_list[index] != ' ':
                    income_total += float(income_list[index])
                else:
                    continue

            for index in range(len(expense_list)):
                if expense_list[index] != ' ':
                    expense_total += float(expense_list[index])
                else:
                    continue

            bookkeeping_particulars += ['今日关帐合计']
            income_list += [format(income_total, '.2f')]
            expense_list += [format(expense_total, '.2f')]
            balance_list += [format(float(value_dict['cmns_tdy_']), '.2f')]

            bookkeeping_df = pd.DataFrame({"描述": bookkeeping_particulars,
                                           "收入": income_list,
                                           "支出": expense_list,
                                           "月累计": balance_list,})
            prtdf(bookkeeping_df)
            print()
            print('——————————————')
            print('正在写入...')
            print()
            promo_df = pd.read_excel(FPara['databaseFileName'], thousands=',', sheet_name=FPara['promoRecordSheetName'])

            tdy_input_promo = [td_pd, dt.datetime.strptime(td_pd,'%Y-%m-%d').strftime("%A"), ]

            for number in range(promotionCountConstant):
                tdy_input_promo += [
                                value_dict['promo{}_ytd_tdy'.format(number+1)][0],
                                value_dict['promo{}_ytd_tdy'.format(number+1)][1]
                ]

            if sec_new:
                promo_updating = {}
                for item in range(len(tdy_input_promo)):
                    promo_updating.update({promo_df.columns[item]:[tdy_input_promo[item]]})

                promo_updating = pd.DataFrame(promo_updating)
                promo_conca = pd.concat([promo_df, promo_updating])
                promo_msg = "{}的促销信息已存入数据库，写入时间: {}".format(take_date, str(pd.to_datetime('today')))

            else:
                promo_conca = promo_df
                promo_msg = "{}的促销信息无法存入或无法再次存入数据库".format(take_date)

            read_net_sales = pd.read_excel(FPara['databaseFileName'], thousands=',', sheet_name=FPara['salesRecordSheetName'])
            tdy_input_sales = [td_pd,
                               dt.datetime.strptime(td_pd,'%Y-%m-%d').strftime("%A"),
                               float(format(float(value_dict['net_sales_aft_deduction']), '.2f')),
                               float(format(float(value_dict['cmns_ytd']), '.2f')),
                               float(format(float(value_dict['cmns_tdy_']), '.2f')),
                               float(value_dict['ads']),]
            if ns_new_input:
                update_sales = {}
                for i in range(len(tdy_input_sales)):
                    update_sales.update({read_net_sales.columns[i]:[tdy_input_sales[i]]})

                update_sales = pd.DataFrame(update_sales)
                conca_sales = pd.concat([read_net_sales, update_sales])
                sales_msg = "{}的营业额信息已存入数据库，写入时间: {}".format(take_date, str(pd.to_datetime('today')))
            else:
                conca_sales = read_net_sales
                sales_msg = "{}的营业额信息无法存入或无法再次存入数据库".format(take_date)

            if cash_record_write:
                read_cash_record = pd.read_excel(FPara['databaseFileName'], thousands=",", sheet_name=FPara['cashRecordSheetName'])
                tdy_cash_record_list = [td_pd, dt.datetime.strptime(td_pd,'%Y-%m-%d').strftime("%A"), cash_received_tdy]

                update_cash_record = {}
                for i in range(len(tdy_cash_record_list)):
                    update_cash_record.update({read_cash_record.columns[i]:[tdy_cash_record_list[i]]})

                update_cash_record = pd.DataFrame(update_cash_record)
                conca_cash_record = pd.concat([read_cash_record, update_cash_record])
                cash_record_msg = "{}的现金信息已存入数据库，写入时间: {}".format(take_date, str(pd.to_datetime('today')))

            else:
                conca_cash_record = read_cash_record
                cash_record_msg = "{}的现金信息无法存入或无法再次存入数据库".format(take_date)


            drink_dict = {
            'drinkTakeOverName0In' : drinkTakeOverName0In,
            'drinkTakeOverName1In' : drinkTakeOverName1In,
            'drinkTakeOverName2In' : drinkTakeOverName2In,
            'drinkTakeOverName3In' : drinkTakeOverName3In,
            'drinkTakeOverName4In' : drinkTakeOverName4In,
            'drinkTakeOverName5In' : drinkTakeOverName5In,
            'drinkTakeOverName6In' : drinkTakeOverName6In,
            'drinkTakeOverName7In' : drinkTakeOverName7In,
            'drinkTakeOverName8In' : drinkTakeOverName8In,
            'drinkTakeOverName9In' : drinkTakeOverName9In,
            'drinkTakeOverName10In' : drinkTakeOverName10In,
            'drinkTakeOverName11In' : drinkTakeOverName11In,

            'drinkTakeOverName0Ziyong' : drinkTakeOverName0Ziyong,
            'drinkTakeOverName1Ziyong' : drinkTakeOverName1Ziyong,
            'drinkTakeOverName2Ziyong' : drinkTakeOverName2Ziyong,
            'drinkTakeOverName3Ziyong' : drinkTakeOverName3Ziyong,
            'drinkTakeOverName4Ziyong' : drinkTakeOverName4Ziyong,
            'drinkTakeOverName5Ziyong' : drinkTakeOverName5Ziyong,
            'drinkTakeOverName6Ziyong' : drinkTakeOverName6Ziyong,
            'drinkTakeOverName7Ziyong' : drinkTakeOverName7Ziyong,
            'drinkTakeOverName8Ziyong' : drinkTakeOverName8Ziyong,
            'drinkTakeOverName9Ziyong' : drinkTakeOverName9Ziyong,
            'drinkTakeOverName10Ziyong' : drinkTakeOverName10Ziyong,
            'drinkTakeOverName11Ziyong' : drinkTakeOverName11Ziyong,

            }

            for number in range(drinkCounter):
                drink_dict.update({ 'drinkTakeOverName{}Out'.format(number) : int(int(value_dict['drinkTakeOverName{}Sale'.format(number)]) + int(drink_dict['drinkTakeOverName{}Ziyong'.format(number)])),
                                    'drinkTakeOverName{}Beizhu'.format(number) : drink_beizhu(drink_dict['drinkTakeOverName{}Ziyong'.format(number)])
                                  })

            read_drink = pd.read_excel(FPara['databaseFileName'], thousands=',', sheet_name=FPara['drinkRecordSheetName'])
            read_drink.sort_values(by='日期', ascending=True, ignore_index=True, inplace=True)

            if len(read_drink.columns) == len(databaseColumns[FPara['drinkRecordSheetName']]):
                if read_drink.columns.tolist() != databaseColumns[FPara['drinkRecordSheetName']]:
                    read_drink.columns = databaseColumns[FPara['drinkRecordSheetName']]
                    print('The columns of {} had been updated.'.format(FPara['drinkRecordSheetName']))

            tdyDrinkRecord = read_drink[read_drink['日期'] == take_date]
            if tdyDrinkRecord.empty:
                if read_drink[read_drink['日期'] == yesterday_date].empty:
                    writeToDrink = False
                else:
                    writeToDrink = True
                    for number in range(drinkCounter):
                        drink_dict.update({ 'drinkTakeOverName{}Previous'.format(number) : int(read_drink[read_drink['日期'] == yesterday_date]['{}实结存'.format(FPara['drinkTakeOverName{}'.format(number)])].values[0]) })

                    for number in range(drinkCounter):
                        drink_dict.update({ 'drinkTakeOverName{}SJC'.format(number) : int(drink_dict['drinkTakeOverName{}Previous'.format(number)]) + int(drink_dict['drinkTakeOverName{}In'.format(number)]) - int(drink_dict['drinkTakeOverName{}Out'.format(number)])})
            else:
                writeToDrink = False

            if writeToDrink:
                out_drink = [take_date]

                if drinkInventoryFunction:
                    for number in range(drinkCounter):
                        out_drink += [
                                        drink_dict['drinkTakeOverName{}Previous'.format(number)],
                                        drink_dict['drinkTakeOverName{}In'.format(number)],
                                        drink_dict['drinkTakeOverName{}Out'.format(number)],
                                        drink_dict['drinkTakeOverName{}SJC'.format(number)],
                                        drink_dict['drinkTakeOverName{}Beizhu'.format(number)]
                                     ]
                else:
                    for number in range(drinkCounter):
                        out_drink += [0,0,0,0,0]

                out_drink_df = {}
                for y in range(len(out_drink)):
                    out_drink_df.update({read_drink.columns[y]:[out_drink[y]]})
                out_drink_df = pd.DataFrame(out_drink_df)
                drink_conca = pd.concat([read_drink, out_drink_df])
                drink_conca.sort_values(by='日期', ascending=True, ignore_index=True, inplace=True)
                drink_msg = "{}的酒水库存已存入数据库，写入时间: {}".format(take_date, str(pd.to_datetime('today')))
            else:
                drink_conca = read_drink
                drink_msg =  "{}的酒水库存无法存入或无法再次存入数据库".format(take_date)

            concaFile = [drink_conca, tabox_concat, conca_sales, conca_cash_record, promo_conca]
            msgList = [drink_msg, tabox_msg, sales_msg, cash_record_msg, promo_msg]
            concaFileName = [FPara['drinkRecordSheetName'],
                             FPara['takeawayBoxRecordSheetName'],
                             FPara['salesRecordSheetName'],
                             FPara['cashRecordSheetName'],
                             FPara['promoRecordSheetName']]

            with pd.ExcelWriter(FPara['databaseFileName']) as writer:
                for file in range(len(concaFile)):
                    concaFile[file].to_excel(writer, sheet_name=concaFileName[file], index=False, header=True, encoding="GBK")

            drink_conca = pd.read_excel(FPara['databaseFileName'], sheet_name=FPara['drinkRecordSheetName'])
            tabox_concat = pd.read_excel(FPara['databaseFileName'], sheet_name=FPara['takeawayBoxRecordSheetName'])
            conca_sales = pd.read_excel(FPara['databaseFileName'], sheet_name=FPara['salesRecordSheetName'])
            conca_cash_record = pd.read_excel(FPara['databaseFileName'], sheet_name=FPara['cashRecordSheetName'])
            promo_conca = pd.read_excel(FPara['databaseFileName'], sheet_name=FPara['promoRecordSheetName'])

            if 'Unnamed: 0' in drink_conca.columns:
                drink_conca.drop(['Unnamed: 0'], axis=1, inplace=True)

            if 'Unnamed: 0' in tabox_concat.columns:
                tabox_concat.drop(['Unnamed: 0'], axis=1, inplace=True)

            if 'Unnamed: 0' in conca_sales.columns:
                conca_sales.drop(['Unnamed: 0'], axis=1, inplace=True)

            if 'Unnamed: 0' in conca_cash_record.columns:
                conca_cash_record.drop(['Unnamed: 0'], axis=1, inplace=True)

            if 'Unnamed: 0' in promo_conca.columns:
                promo_conca.drop(['Unnamed: 0'], axis=1, inplace=True)

            with pd.ExcelWriter(FPara['databaseFileName']) as writer:
                for file in range(len(concaFile)):
                    concaFile[file].to_excel(writer, sheet_name=concaFileName[file], index=False, header=True, encoding="GBK")
                    print(msgList[file])

            print()
            print('——————————————')
            NameConverter = readCsv(githubFileName=NameConverterForShowingDfCsvName,
                           githubUserName=FPara['githubUserName'],
                           githubRepoName=FPara['githubRepoName'],
                           githubBranchName=FPara['githubBranchName'],
                           csvSep='|',
                           csvEncoding='utf-8',
                           runLocally=FPara['runLocally'])

            if drinkInventoryFunction:
                print()
                print('酒水信息: ')
                print(take_date)
                read_drink = pd.read_excel(FPara['databaseFileName'], thousands = ',', sheet_name=FPara['drinkRecordSheetName'])
                for_drink_df = read_drink[read_drink['日期'] == take_date]

                yinLiao = []
                drinkRuku = []
                drinkShouchu = []
                drinkZiyong = []
                drinkShengyu = []

                for number in range(drinkCounter):
                    yinLiao += [NameConverter[NameConverter['name in rule'] == 'drinkTakeOverName{}'.format(number)]['name to show'].values[0]]

                    drinkRuku += [for_drink_df['{}进'.format(FPara['drinkTakeOverName{}'.format(number)])].values[0]]

                    drinkShouchu += [int(value_dict['drinkTakeOverName{}Sale'.format(number)])]

                    drinkZiyong += [int(for_drink_df['{}出'.format(FPara['drinkTakeOverName{}'.format(number)])].values[0]) - int(value_dict['drinkTakeOverName{}Sale'.format(number)])]

                    drinkShengyu += [int(for_drink_df['{}实结存'.format(FPara['drinkTakeOverName{}'.format(number)])].values[0])]


                drink_df = pd.DataFrame({'饮料(单位)' : yinLiao, '入库' : drinkRuku, '售出' : drinkShouchu, '自用' : drinkZiyong, '剩余': drinkShengyu})
                drink_df['售出'] = drink_df['售出'].apply(lambda x : -1*x)
                drink_df['自用'] = drink_df['自用'].apply(lambda x : -1*x)

                dropIndex = []
                for index in range(len(drink_df)):
                    if str(drink_df.iloc[index, 0]).startswith("饮料"):
                        dropIndex += [index]

                drink_df.drop(dropIndex, axis=0, inplace=True)
                prtdf(drink_df)
            else:
                pass

            if takeawayBoxInventoryFunction:
                print()
                print('打包盒信息: ')
                print(take_date)

                tabox = pd.read_excel(FPara['databaseFileName'], sheet_name=FPara['takeawayBoxRecordSheetName'])
                for_takeaway_df = tabox[tabox['Date'] == td_pd]
                wuPin = []
                ruKu = []
                yongChu = []
                shengYu = []

                for number in range(boxCounter):
                    wuPin += [NameConverter[NameConverter['name in rule'] == 'boxTakeOverName{}'.format(number)]['name to show'].values[0]]

                    ruKu += [for_takeaway_df['{}入库'.format(FPara['boxTakeOverName{}'.format(number)])].values[0]]

                    yongChu += [for_takeaway_df['{}出库'.format(FPara['boxTakeOverName{}'.format(number)])].values[0]]

                    shengYu += [for_takeaway_df['{}现有数量'.format(FPara['boxTakeOverName{}'.format(number)])].values[0]]

                takeaway_df = pd.DataFrame({'物品(单位)': wuPin, '入库': ruKu, '用出': yongChu, '剩余' : shengYu})
                takeaway_df['用出'] = takeaway_df['用出'].apply(lambda x:-1*round(x,2))
                takeaway_df['剩余'] = takeaway_df['剩余'].apply(lambda x:round(x,2))

                dropIndex = []
                for index in range(len(takeaway_df)):
                    if str(takeaway_df.iloc[index, 0]).startswith("box"):
                        dropIndex += [index]

                takeaway_df.drop(dropIndex, axis=0, inplace=True)
                prtdf(takeaway_df)
                print()

            else:
                pass

            if cashRecordShowFunction:
                print("现金收入存放信息: ")
                print(take_date)
                cash_df = pd.read_excel(FPara['databaseFileName'], thousands=",", sheet_name=FPara['cashRecordSheetName'])
                cash_df['Date'] = cash_df['Date'].astype(str)
                if cash_received_tdy != None:
                    start_date_continue = True
                    try:
                        start_date_index = int(cash_df[cash_df['Date'] == cash_start_date].index.values[0])
                    except:
                        print("未存放现金的起始日期错误,原因可能是标点符号使用错误或是数据库里不存在该日期的记录. ")
                        start_date_continue = False

                    if start_date_continue:
                        if cash_end_date == 'default':
                            if start_date_index != 0:
                                drop_indexes = np.arange(start_date_index)
                                cash_df.drop(drop_indexes, axis=0, inplace=True)
                                cash_df.reset_index(inplace=True)
                                if "index" in  cash_df.columns:
                                    cash_df.drop("index", axis=1, inplace=True)
                            else:
                                pass

                            cash_summary = cash_df['Amount'].sum()
                            prtdf(cash_df)
                            print("{}至{}的现金收入共计${}".format(cash_start_date, td_pd, format(cash_summary, '.2f')))

                        else:
                            end_date_continue = True
                            try:
                                end_date_index = int(cash_df[cash_df['Date'] == cash_end_date].index.values[0])
                            except:
                                print("未存放现金的结束日期错误,原因可能是标点符号使用错误或是数据库里不存在该日期的记录. ")
                                end_date_continue = False

                            if end_date_continue:
                                prtdf(cash_df.iloc[np.arange(start_date_index, end_date_index+1),:])
                                cash_summary = cash_df.iloc[np.arange(start_date_index, end_date_index+1),:]["Amount"].sum()
                                print("{}至{}的现金收入共计${}".format(cash_start_date, cash_end_date, format(cash_summary, '.2f')))
                            else:
                                pass
                    else:
                        pass
                else:
                    print("现金收入存放信息生成错误! ")
            else:
                pass
        else:
             print("逻辑错误")
    else:
        print("导出的报表无法识别来自哪家门店")
        print("请选择门店导出报表后再运行一遍吧")
else:
    print("你应该没有复制所有的后台数据,复制完所有的数据再允许一遍吧")
