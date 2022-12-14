#! /usr/bin/env python3

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
from requests.structures import CaseInsensitiveDict
import io
import sys
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
    mleft = df[df['Date'] == date]['{}????????????'.format(df_string)].values[0]
    return mleft

def model_now(df, key_name, box_dict, td_ytd):
    if float(box_dict['{}_absolute'.format(key_name)]) < 0:
        xianYouShuLiang = float(df[df['Date'] == td_ytd]['{}????????????'.format(key_name)].values[0])
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
                    '{}'.format(model_df.iloc[index, 0]) : box_value['{}_now'.format(model_df.iloc[index, 0])]
            })

    elif not alert_everyday:
        if day_name == 'Friday':
            for index in range(len(model_df)):
                alert_dict.update({
                        '{}'.format(model_df.iloc[index, 0]) : box_value['{}_now'.format(model_df.iloc[index, 0])]
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
                    '{}'.format(noti_df.iloc[index, 0]) : '??????????????????{}'.format(noti_df.iloc[index, 1])
            })

        else:
            continue

    return [alert_dict, noti_alert]

def sending_email(text, receiver_list, td_pd, emailSender, emailSenderPassword):
    mail_server = 'smtp.126.com'

    message = MIMEText(str(text), 'plain', 'utf-8')
    message['From'] = emailSender

    subject = '?????????????????? {}'.format(td_pd)
    message['Subject'] = Header(subject, 'utf-8')

    try:
        smtpObj = smtplib.SMTP()
        smtpObj.connect(mail_server, 25)
        smtpObj.login(emailSender, emailSenderPassword)
        for ppl in receiver_list:
            print('?????????????????????{}'.format(ppl))
            message['To'] = ppl
            smtpObj.sendmail(emailSender, ppl, message.as_string())
            print('{} ??????????????????'.format(ppl))
    except smtplib.SMTPException:
        print('??????????????????')

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
        return str(ziyong) + '???/?????????'

def prtdf(df):
    with pd.option_context('display.max_rows', None,
                           'display.max_columns', None,
                           'display.width', 1000,
                           'display.precision', 2,
                           'display.colheader_justify', 'center'):
        return display(df)

def readCsv(githubUserName, githubRepoName, githubBranchName, githubFileName, githubFolderName, csvSep, csvEncoding, runLocally):
    if runLocally:
        return pd.read_csv("{}/{}/{}".format(os.getcwd(), githubFolderName, githubFileName), sep=csvSep, encoding=csvEncoding)
    else:
        githubPrefix = "https://raw.githubusercontent.com"
        URL = "{}/{}/{}/{}/{}/{}".format(githubPrefix, githubUserName, githubRepoName, githubBranchName, githubFolderName, githubFileName)

        resp = requests.get(URL)
        if resp.status_code == 200:
            return pd.read_csv(io.StringIO(resp.text), sep=csvSep, encoding=csvEncoding)
        else:
            raise ValueErorr("????????????")

#???????????????
#lun_sales = 0

#??????????????????
#lun_gc = 0

#????????????????????????
#lun_fwc = '4'

#????????????????????????
#lun_kwc = '7+1?????????'

#??????????????????
#tb_sales = 0

#???????????????
#tb_gc = 0

#???????????????????????????
#tb_fwc = '1'

#???????????????????????????
#tb_kwc = '0'

#????????????????????????
#night_fwc = '4'

#????????????????????????
#night_kwc = '7+1?????????'


#--------???????????????------------#

#??????????????????
#coke_in = 0
#??????????????????
#coke_ziyong = 0

#??????????????????
#tiger_in = 0
#??????????????????
#tiger_ziyong = 0

#??????????????????
#heineken_in = 0
#??????????????????
#heineken_ziyong = 0

#?????????????????????
#water_in = 0
#?????????????????????
#water_ziyong = 0

#??????????????????
#redwine_in = 0
#??????????????????
#redwine_ziyong = 0

#??????????????????
#whitewine_in = 0
#??????????????????
#whitewine_ziyong = 0

#??????????????????????????????
#xmas_wine_in = 0
#??????????????????????????????
#xmas_wine_ziyong = 0

#--------?????????????????????----------#

#??????????????????????????????(??????: ????????????-??????-??????)
#????????????????????????????????????
#cash_start_date = '2022-11-01'

#??????????????????????????????(??????: ????????????-??????-??????, ?????????: 'default')
#'default'??????????????????????????????????????????????????????????????????
#????????????????????????????????????
#cash_end_date = 'default'

#--------??????book?????????----------#
#??????????????????????????????????????????????????????,????????????????????????????????????
#???????????????????????????????????????,???theOutlet
#??????
#theOutlet = "Thomson"

#githubUserName = ""
#githubRepoName = ""
#githubBranchName = ""
#githubFolderName = ""
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

#print("??????????????????...")
#on_net = on_internet()

#if not on_net:
#    print('??????????????????')
#else:
#    print("??????????????????")

#URL = "{}/{}/{}/{}/{}".format("https://raw.githubusercontent.com", githubUserName, githubRepoName, githubBranchName, githubFileName)
#script = urllib.request.urlopen(URL).read().decode()
#exec(script)

functionRuleDf = readCsv(githubFileName=functionalityRuleCsvName,
                         githubUserName=githubUserName,
                         githubRepoName=githubRepoName,
                         githubBranchName=githubBranchName,
                         githubFolderName=githubFolderName,
                         csvSep=',',
                         csvEncoding='utf-8',
                         runLocally=runLocally)

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
FPara['githubFolderName'] = githubFolderName
FPara['takeawayBoxInventoryURL'] = takeawayBoxInventoryURL
FPara['runLocally'] = runLocally

takeawayBoxInventoryFunction = FPara['takeawayBoxInventoryFunction']
stockAlertFunction = FPara['stockAlertFunction']
stockEverydayFunction = FPara['stockEverydayFunction']
cashRecordShowFunction = FPara['cashRecordShowFunction']
drinkInventoryFunction = FPara['drinkInventoryFunction']


chineseFunction = ["?????????????????????", "?????????????????????????????????",
                   "???????????????????????????????????????",
                   "??????????????????????????????", "??????????????????",]
functionVarList = [takeawayBoxInventoryFunction,
                   stockAlertFunction,
                   stockEverydayFunction,
                   cashRecordShowFunction,
                   drinkInventoryFunction]

print()
for index in range(len(chineseFunction)):
    if functionVarList[index]:
        print("{}?????????".format(chineseFunction[index]))
    else:
        print("{}?????????".format(chineseFunction[index]))

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

drinkRecordColumns = ['??????']
for num in range(drinkCounter):
    drinkRecordColumns += [
                           '{}????????????'.format(FPara['drinkTakeOverName{}'.format(num)]),
                           '{}???'.format(FPara['drinkTakeOverName{}'.format(num)]),
                           '{}???'.format(FPara['drinkTakeOverName{}'.format(num)]),
                           '{}?????????'.format(FPara['drinkTakeOverName{}'.format(num)]),
                           '{}??????'.format(FPara['drinkTakeOverName{}'.format(num)])
                          ]


databaseColumns[FPara['drinkRecordSheetName']] = drinkRecordColumns

boxRecordColumns = ['Date']
for num in range(boxCounter):
    boxRecordColumns += [
                         '{}??????'.format(FPara['boxTakeOverName{}'.format(num)]),
                         '{}??????'.format(FPara['boxTakeOverName{}'.format(num)]),
                         '{}????????????'.format(FPara['boxTakeOverName{}'.format(num)])

                        ]

databaseColumns[FPara['takeawayBoxRecordSheetName']] = boxRecordColumns

read = pd.read_excel(FPara['bookFileName'],thousands=',')
columnsRename = np.arange(0, len(read.columns)).astype(str)
read.columns = columnsRename
read['0'] = read['0'].astype(str)

if len(read) > int(FPara['minBookFileLenAllowable']):
    isLocalBook = not read[read["0"].str.contains("GST Reg No")].empty

    if isLocalBook:
        try:
            outlet_loc = theOutlet
            continue_ = True
        except NameError:
            continue_ = False

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

    else:
        try:
            bi = int(read[read['0'].str.contains('Date:')].index.values[0])
            ei = int(read[read['0'] == 'Total Sales'].index.values[0])
            outlet_loc = str(read.iloc[(bi+1):(ei),:].dropna(axis=1).values[0][0])
            continue_ = True
        except IndexError:
            continue_ = False

    if continue_:
        pbi = int(read[read['0'].str.contains('PAYMENT BREAKDOWN:')].index[0])

        if isLocalBook:
            pei = int(read[read['0'].str.contains('OPENING CASH BALANCE:')].index[0])

        else:
            try:
                pei = int(read[read['0'].str.contains('PAYMENT BREAKDOWN (POS):', regex=False)].index[0])
            except IndexError:
                pei = int(read[read['0'].str.contains('Transaction Void Items')].index[0])

        pay_breakdown = read.copy()
        pay_breakdown = pay_breakdown.iloc[(pbi+1):(pei), :]
        resetAxis(pay_breakdown, axis=0)
        pay_breakdown['0'] = pay_breakdown['0'].astype(str)

        if isLocalBook:
            drop_index = min(pay_breakdown[pay_breakdown['0'].str.contains('nan')].index)
            pay_breakdown = pay_breakdown.iloc[:drop_index]

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
            print("?????????????????????")

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
        if sys.platform.strip().upper() == "IOS":
            #it may not render strftime with Chinese characters correctly on ios app such as Juno
            today_date = today_dt_format.strftime("%YNIAN%mYUE%RI")
            today_date = today_date.replace("NIAN", "???")
            today_date = today_date.replace("YUE", "???")
            today_date = today_date.replace("RI", "???")
        else:
            today_date = today_dt_format.strftime("%Y???%m???%d???")

        #raw_date_from_book
        if isLocalBook:

            try:
                read_date = read.iloc[int(read[read["0"].str.contains('X/Shift Report')].index[0])+1,0]
            except IndexError:
                read_date = read.iloc[int(read[read["0"].str.contains('Z/Closing Report')].index[0])+1,0]

            rdfb = read_date.split()

        else:
            rdfb = read[read['0'].str.contains('Date:')].dropna(axis=1).values[0][0]
            rdfb = rdfb.split()

        if isLocalBook:
            dfb = '{} {} {}'.format(rdfb[1], rdfb[2], rdfb[3])
            dfb = dt.datetime.strptime(dfb, '%d %b %Y')

        else:
            if len(rdfb) == 8:
                if rdfb[1] == rdfb[5] and rdfb[2] == rdfb[6] and rdfb[3] == rdfb[7]:
                    #date_from_book
                    dfb = '{} {} {}'.format(rdfb[1],rdfb[2],rdfb[3])
                    dfb = dt.datetime.strptime(dfb, '%d %b %Y')

                else:
                    print('?????????????????????\n??????????????????????????????????????????')
                    dfb = today_dt_format
            else:
                print('????????????????????????????????????????????????????????????')
                dfb = today_dt_format

        if sys.platform.strip().upper() == "IOS":
            #it may not render strftime with Chinese characters correctly on ios app such as Juno
            take_date = dfb.strftime("%YNIAN%mYUE%dRI")
            take_date = take_date.replace("NIAN", "???")
            take_date = take_date.replace("YUE", "???")
            take_date = take_date.replace("RI", "???")
        else:
            take_date = dfb.strftime("%Y???%m???%d???")

        td_pd = dfb.strftime('%Y-%m-%d')
        td_ytd = (dfb - dt.timedelta(days=1)).strftime('%Y-%m-%d')

        if sys.platform.strip().upper() == "IOS":
            yesterday_date = (dfb - dt.timedelta(days=1)).strftime('%YNIAN%mYUE%dRI')
            yesterday_date = yesterday_date.replace("NIAN", "???")
            yesterday_date = yesterday_date.replace("YUE", "???")
            yesterday_date = yesterday_date.replace("RI", "???")
        else:
            yesterday_date = (dfb - dt.timedelta(days=1)).strftime('%Y???%m???%d???')

        TBRuleDf = readCsv(githubFileName=takeawayBoxRuleCsvName,
                       githubUserName=FPara['githubUserName'],
                       githubRepoName=FPara['githubRepoName'],
                       githubBranchName=FPara['githubBranchName'],
                       githubFolderName=FPara['githubFolderName'],
                       csvSep=',',
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
                           githubFolderName=FPara['githubFolderName'],
                           csvSep=',',
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
            tabox_msg = "{}???????????????????????????????????????????????????: {}".format(take_date, str(pd.to_datetime('today')))

            if stockAlertFunction:
                if takeawayBoxInventoryFunction:
                    alerts = stock_alert(box_value=box_value, model_df=model_df, noti_df=noti_df, td_pd=td_pd, alert_everyday=stockEverydayFunction)
                    if len(alerts[0]) == 0 and len(alerts[1]) == 0:
                        send_email = False
                    elif len(alerts[0]) != 0 and len(alerts[1]) == 0:
                        send_email = True
                        text = '?????????: \n'
                        for key, items in alerts[0].items():
                            text += '{}: ???{} (???/???) \n'.format(key, round(float(items), 3))
                            text += '\n'

                    elif len(alerts[0]) == 0 and len(alerts[1]) != 0:
                        send_email = True
                        text = '{}?????????????????????: \n'.format(td_pd)
                        for key, items in alerts[1].items():
                            text += '{}'.format(key)
                            text += '\n'

                    elif len(alerts[0]) != 0 and len(alerts[1]) != 0:
                        send_email=True

                        text = '?????????: \n'
                        for key, items in alerts[0].items():
                            text += '{}: ???{} (???/???) \n'.format(key, round(float(items), 3))
                            text += '\n'

                        text += '\n'

                        text += '{}?????????????????????: \n'.format(td_pd)
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
            tabox_msg = "{}????????????????????????????????????????????????????????????".format(take_date)

        print("?????????...???????????????")
        total_sales = float(resetAxis(read[read['0'] == 'Total Sales'].dropna(axis=1), axis=1)['1'].values[0])
        no_of_cover = int(resetAxis(read[read['0'] == 'No. Of Cover'].dropna(axis=1), axis=1)['1'].values[0])

        if isLocalBook:
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
                lun_avg = lun_sales/lun_gc

            lun_avg = normal_round(lun_avg, 2)

        if isLocalBook:
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
                tb_avg = tb_sales/tb_gc

            tb_avg = normal_round(tb_avg, 2)

        if isLocalBook:
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
            night_gc = no_of_cover - lun_gc - tb_gc

            if night_gc == 0:
                night_avg = 0
            else:
                night_avg = night_sales/night_gc

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
                       githubFolderName=FPara['githubFolderName'],
                       csvSep=',',
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
        value_dict['net_sales_aft_deduction'] = single_zero(format(normal_round(value_dict['net_sales_aft_deduction'], 2), ".2f"))

        read_net_sales = pd.read_excel(FPara['databaseFileName'], thousands=',', sheet_name=FPara['salesRecordSheetName'])
        dfb_record_exist = read_net_sales[read_net_sales['Date'] == td_pd]
        if dfb_record_exist.empty:
            if read_net_sales['Date'].values[-1] == td_ytd:
                if int(dfb.day) == 1:
                    cmns_ytd = 0.0
                    cmns_tdy = cmns_ytd + float(value_dict['net_sales_aft_deduction'])
                    cmns_tdy_ = format(normal_round(cmns_tdy, 2), ".2f")
                    avg_daily_sales = cmns_tdy/int(dfb.day)
                    ads_rd1 = str(avg_daily_sales)[str(avg_daily_sales).find('.'):][:3]
                    ads_rd2 = str(avg_daily_sales)[:str(avg_daily_sales).find('.')]
                    ads = float(ads_rd2+ads_rd1)
                    ns_new_input = True
                else:
                    cmns_ytd = float(read_net_sales['cum net sales today'].values[-1])
                    cmns_tdy = cmns_ytd + float(value_dict['net_sales_aft_deduction'])
                    cmns_tdy_ = format(normal_round(cmns_tdy, 2),".2f")
                    avg_daily_sales = cmns_tdy/int(dfb.day)
                    ads_rd1 = str(avg_daily_sales)[str(avg_daily_sales).find('.'):][:3]
                    ads_rd2 = str(avg_daily_sales)[:str(avg_daily_sales).find('.')]
                    ads = float(ads_rd2+ads_rd1)
                    ns_new_input = True
            else:
                print('{}?????????({})????????????????????????!'.format(td_pd, td_ytd))
                cmns_ytd = np.nan
                cmns_tdy_ = np.nan
                ads = np.nan
                ns_new_input = False
        else:
            print('{}????????????????????????????????????'.format(td_pd))
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
                print('{}?????????({})?????????????????????!'.format(td_pd, td_ytd))
                sec_new = False
        else:
            print('{}?????????????????????????????????'.format(td_pd))
            sec_new = False

        read_cash_record = pd.read_excel(FPara['databaseFileName'], thousands=",", sheet_name=FPara['cashRecordSheetName'])
        rcr_record_exist = read_cash_record[read_cash_record["Date"] == td_pd]
        if rcr_record_exist.empty:
            if read_cash_record["Date"].values[-1] == td_ytd:
                cash_received_tdy = float(value_dict["CASH"])
                cash_record_write = True
            else:
                print('{}?????????({})?????????????????????!'.format(td_pd, td_ytd))
                cash_received_tdy = None
                cash_record_write = False
        else:
            print("{}?????????????????????????????????".format(td_pd))
            cash_received_tdy = float(read_cash_record[read_cash_record["Date"] == td_pd]["Amount"].values[0])
            cash_record_write = False

        generate=True
        for num in range(promotionCountConstant):
            if int(value_dict['promo{}_ytd_tdy'.format(num+1)][1]) < int(value_dict['promo{}_count'.format(num+1)]):
                generate = False
            else:
                pass
        print()
        print('???????????????...')
        print('??????????????????????????????????????????')
        print()
        if generate:
            print_result = []
            printRuleDf = readCsv(githubFileName=printRuleCsvName,
                           githubUserName=FPara['githubUserName'],
                           githubRepoName=FPara['githubRepoName'],
                           githubBranchName=FPara['githubBranchName'],
                           githubFolderName=FPara['githubFolderName'],
                           csvSep=',',
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

            bookkeeping_particulars = ['?????????', '?????????']

            voucher_expense_list = ['Trueblue??????', '???????????????', 'PandaBox', "??????????????????$5"]
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

            bookkeeping_particulars += ['??????????????????']
            income_list += [format(income_total, '.2f')]
            expense_list += [format(expense_total, '.2f')]
            balance_list += [format(float(value_dict['cmns_tdy_']), '.2f')]

            bookkeeping_df = pd.DataFrame({"??????": bookkeeping_particulars,
                                           "??????": income_list,
                                           "??????": expense_list,
                                           "?????????": balance_list,})

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
                promo_msg = "{}????????????????????????????????????????????????: {}".format(take_date, str(pd.to_datetime('today')))

            else:
                promo_conca = promo_df
                promo_msg = "{}?????????????????????????????????????????????????????????".format(take_date)

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
                sales_msg = "{}???????????????????????????????????????????????????: {}".format(take_date, str(pd.to_datetime('today')))
            else:
                conca_sales = read_net_sales
                sales_msg = "{}????????????????????????????????????????????????????????????".format(take_date)

            if cash_record_write:
                read_cash_record = pd.read_excel(FPara['databaseFileName'], thousands=",", sheet_name=FPara['cashRecordSheetName'])
                tdy_cash_record_list = [td_pd, dt.datetime.strptime(td_pd,'%Y-%m-%d').strftime("%A"), cash_received_tdy]

                update_cash_record = {}
                for i in range(len(tdy_cash_record_list)):
                    update_cash_record.update({read_cash_record.columns[i]:[tdy_cash_record_list[i]]})

                update_cash_record = pd.DataFrame(update_cash_record)
                conca_cash_record = pd.concat([read_cash_record, update_cash_record])
                cash_record_msg = "{}????????????????????????????????????????????????: {}".format(take_date, str(pd.to_datetime('today')))

            else:
                conca_cash_record = read_cash_record
                cash_record_msg = "{}?????????????????????????????????????????????????????????".format(take_date)


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
            read_drink.sort_values(by='??????', ascending=True, ignore_index=True, inplace=True)

            if len(read_drink.columns) == len(databaseColumns[FPara['drinkRecordSheetName']]):
                if read_drink.columns.tolist() != databaseColumns[FPara['drinkRecordSheetName']]:
                    read_drink.columns = databaseColumns[FPara['drinkRecordSheetName']]
                    print('The columns of {} had been updated.'.format(FPara['drinkRecordSheetName']))

            tdyDrinkRecord = read_drink[read_drink['??????'] == take_date]
            if tdyDrinkRecord.empty:
                if read_drink[read_drink['??????'] == yesterday_date].empty:
                    writeToDrink = False
                else:
                    writeToDrink = True
                    for number in range(drinkCounter):
                        drink_dict.update({ 'drinkTakeOverName{}Previous'.format(number) : int(read_drink[read_drink['??????'] == yesterday_date]['{}?????????'.format(FPara['drinkTakeOverName{}'.format(number)])].values[0]) })

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
                drink_conca.sort_values(by='??????', ascending=True, ignore_index=True, inplace=True)
                drink_msg = "{}????????????????????????????????????????????????: {}".format(take_date, str(pd.to_datetime('today')))
            else:
                drink_conca = read_drink
                drink_msg =  "{}?????????????????????????????????????????????????????????".format(take_date)

            concaFile = [drink_conca, tabox_concat, conca_sales, conca_cash_record, promo_conca]
            msgList = [drink_msg, tabox_msg, sales_msg, cash_record_msg, promo_msg]
            concaFileName = [FPara['drinkRecordSheetName'],
                             FPara['takeawayBoxRecordSheetName'],
                             FPara['salesRecordSheetName'],
                             FPara['cashRecordSheetName'],
                             FPara['promoRecordSheetName']]

            with pd.ExcelWriter(FPara['databaseFileName']) as writer:
                for file in range(len(concaFile)):
                    concaFile[file].to_excel(writer, sheet_name=concaFileName[file], index=False, header=True)

            drink_conca = pd.read_excel(FPara['databaseFileName'], sheet_name=FPara['drinkRecordSheetName'])
            tabox_concat = pd.read_excel(FPara['databaseFileName'], sheet_name=FPara['takeawayBoxRecordSheetName'])
            conca_sales = pd.read_excel(FPara['databaseFileName'], sheet_name=FPara['salesRecordSheetName'])
            conca_cash_record = pd.read_excel(FPara['databaseFileName'], sheet_name=FPara['cashRecordSheetName'])
            promo_conca = pd.read_excel(FPara['databaseFileName'], sheet_name=FPara['promoRecordSheetName'])

            NameConverter = readCsv(githubFileName=NameConverterForShowingDfCsvName,
                           githubUserName=FPara['githubUserName'],
                           githubRepoName=FPara['githubRepoName'],
                           githubBranchName=FPara['githubBranchName'],
                           githubFolderName=FPara['githubFolderName'],
                           csvSep=',',
                           csvEncoding='utf-8',
                           runLocally=FPara['runLocally'])

            if drinkInventoryFunction:
                read_drink = pd.read_excel(FPara['databaseFileName'], thousands = ',', sheet_name=FPara['drinkRecordSheetName'])
                for_drink_df = read_drink[read_drink['??????'] == take_date]

                yinLiao = []
                drinkRuku = []
                drinkShouchu = []
                drinkZiyong = []
                drinkShengyu = []

                for number in range(drinkCounter):
                    yinLiao += [NameConverter[NameConverter['name in rule'] == 'drinkTakeOverName{}'.format(number)]['name to show'].values[0]]

                    drinkRuku += [for_drink_df['{}???'.format(FPara['drinkTakeOverName{}'.format(number)])].values[0]]

                    drinkShouchu += [int(value_dict['drinkTakeOverName{}Sale'.format(number)])]

                    drinkZiyong += [int(for_drink_df['{}???'.format(FPara['drinkTakeOverName{}'.format(number)])].values[0]) - int(value_dict['drinkTakeOverName{}Sale'.format(number)])]

                    drinkShengyu += [int(for_drink_df['{}?????????'.format(FPara['drinkTakeOverName{}'.format(number)])].values[0])]


                drink_df = pd.DataFrame({'??????(??????)' : yinLiao, '??????' : drinkRuku, '??????' : drinkShouchu, '??????' : drinkZiyong, '??????': drinkShengyu})
                drink_df['??????'] = drink_df['??????'].apply(lambda x : -1*x)
                drink_df['??????'] = drink_df['??????'].apply(lambda x : -1*x)

                dropIndex = []
                for index in range(len(drink_df)):
                    if str(drink_df.iloc[index, 0]).startswith("??????"):
                        dropIndex += [index]

                drink_df.drop(dropIndex, axis=0, inplace=True)

            else:
                drink_df = None

            if takeawayBoxInventoryFunction:
                tabox = pd.read_excel(FPara['databaseFileName'], sheet_name=FPara['takeawayBoxRecordSheetName'])
                for_takeaway_df = tabox[tabox['Date'] == td_pd]
                wuPin = []
                ruKu = []
                yongChu = []
                shengYu = []

                for number in range(boxCounter):
                    wuPin += [NameConverter[NameConverter['name in rule'] == 'boxTakeOverName{}'.format(number)]['name to show'].values[0]]

                    ruKu += [for_takeaway_df['{}??????'.format(FPara['boxTakeOverName{}'.format(number)])].values[0]]

                    yongChu += [for_takeaway_df['{}??????'.format(FPara['boxTakeOverName{}'.format(number)])].values[0]]

                    shengYu += [for_takeaway_df['{}????????????'.format(FPara['boxTakeOverName{}'.format(number)])].values[0]]

                takeaway_df = pd.DataFrame({'??????(??????)': wuPin, '??????': ruKu, '??????': yongChu, '??????' : shengYu})
                takeaway_df['??????'] = takeaway_df['??????'].apply(lambda x:-1*round(x,2))
                takeaway_df['??????'] = takeaway_df['??????'].apply(lambda x:round(x,2))

                dropIndex = []
                for index in range(len(takeaway_df)):
                    if str(takeaway_df.iloc[index, 0]).startswith("box"):
                        dropIndex += [index]

                takeaway_df.drop(dropIndex, axis=0, inplace=True)


            else:
                takeaway_df = None

            if cashRecordShowFunction:

                cash_df = pd.read_excel(FPara['databaseFileName'], thousands=",", sheet_name=FPara['cashRecordSheetName'])
                cash_df['Date'] = cash_df['Date'].astype(str)
                if cash_received_tdy != None:
                    start_date_continue = True
                    try:
                        start_date_index = int(cash_df[cash_df['Date'] == cash_start_date].index.values[0])
                    except:
                        cash_df = None
                        cash_summary = None
                        cash_df_msg = "????????????????????????????????????,????????????????????????????????????????????????????????????????????????????????????. "
                        start_date_continue = False

                    if start_date_continue:
                        if cash_end_date == 'default':
                            if start_date_index != 0:
                                drop_indexes = np.arange(start_date_index)
                                cash_df.drop(drop_indexes, axis=0, inplace=True)
                                cash_df.reset_index(inplace=True)
                                if "index" in cash_df.columns:
                                    cash_df.drop("index", axis=1, inplace=True)
                            else:
                                pass

                            cash_summary = cash_df['Amount'].sum()
                            cash_df_msg = "{}???{}?????????????????????${}".format(cash_start_date, td_pd, format(cash_summary, '.2f'))

                        else:
                            end_date_continue = True
                            try:
                                end_date_index = int(cash_df[cash_df['Date'] == cash_end_date].index.values[0])
                            except:
                                cash_df = None
                                cash_summary = None
                                cash_df_msg = "????????????????????????????????????,????????????????????????????????????????????????????????????????????????????????????. "
                                end_date_continue = False

                            if end_date_continue:
                                cash_df = cash_df.iloc[np.arange(start_date_index, end_date_index+1),:]
                                cash_summary = cash_df["Amount"].sum()
                                cash_df_msg = "{}???{}?????????????????????${}".format(cash_start_date, cash_end_date, format(cash_summary, '.2f'))
                            else:
                                pass
                    else:
                        pass
                else:
                    cash_df = None
                    cash_summary = None
                    cash_df_msg = "????????????????????????????????????! "
            else:
                pass


            for items in print_result:
                print(items)

            print()
            print('??????????????????????????????????????????')
            print()
            print('?????????: ${}'.format(value_dict['svc']))
            print('GST: ${}'.format(value_dict['gst']))
            print('???????????????: ${}'.format(value_dict['ads']))
            print()
            print("????????????")
            print(take_date)
            print()
            prtdf(bookkeeping_df)
            print()
            print('??????????????????????????????????????????')
            print('????????????...')
            for msg in msgList:
                print(msg)
            print()
            print('??????????????????????????????????????????')
            if drinkInventoryFunction:
                if isinstance(drink_df, pd.DataFrame):
                    print('????????????: ')
                    print(take_date)
                    prtdf(drink_df)
                    print()

            if takeawayBoxInventoryFunction:
                if isinstance(takeaway_df, pd.DataFrame):
                    print('???????????????: ')
                    print(take_date)
                    prtdf(takeaway_df)
                    print()

            if cashRecordShowFunction:
                print("????????????????????????: ")
                print(take_date)
                if isinstance(cash_df, pd.DataFrame) and cash_summary != None:
                    prtdf(cash_df)
                    print(cash_df_msg)
                else:
                    print(cash_df_msg)

        else:
             print("????????????")
    else:
        print("?????????????????????????????????????????????")
        print("????????????????????????????????????????????????")
        print("????????????????????????????????????,?????????theOutlet??????????????????")
else:
    print("??????????????????????????????????????????,??????????????????????????????????????????")
