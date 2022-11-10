#! /usr/bin/env python3
#encoding='GBK'

import numpy as np
import pandas as pd
import datetime as dt
import requests
import re
import os
from requests.structures import CaseInsensitiveDict
import io
#import urllib

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

def prtdf(df):
    with pd.option_context('display.max_rows', None,
                           'display.max_columns', None,
                           'display.width', 1000,
                           'display.precision', 2,
                           'display.colheader_justify', 'center'):
        return display(df)

def DfRemoveUnits(regex, original):
    reStr = re.search(regex, original)
    if reStr is None:
        return reStr
    else:
        return original.replace(reStr[0], '')

def integer_check(x):
    try:
        x = int(x)
        return True
    except ValueError:
        return False

def need_integer(x, absolute=False):
    if integer_check(x) and absolute:
        return abs(int(x))

    elif integer_check(x) and not absolute:
        return int(x)

    else:
        while not integer_check(x):
            print('输入的'+ str(x) + '不是整数!')
            x = input('请输入整数: ')
            integer_check(x)
            while integer_check(x):
                if absolute:
                    return abs(int(x))
                else:
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

#def on_internet():
#    url = "https://www.google.com.sg/"
#    timeout = 5
#    try:
#        request = requests.get(url, timeout=timeout)
#        return True
#    except (requests.ConnectionError, requests.Timeout) as exception:
#        return False

def DocFileNameCheck(fileName):
    regexYear = r'.*\年'
    regexMonth = r'年.*月'
    if re.search(regexYear, fileName) is None:
        raise TypeError
    else:
        if re.search(regexMonth, fileName) is None:
            raise TypeError
        else:
            return [re.search(regexYear, fileName)[0].replace("年", ""), re.search(regexMonth, fileName)[0].replace("月", "").replace("年", "")]

def gen_pdf(customFont, pageOneDf, pageTwoDf, pageThreeDf, fileNameAfterCheck, outlet):
    regexYear = r'.*\年'

    Document = borb_Document()
    Page = borb_Page(width=Decimal(595), height=Decimal(842))
    Document.append_page(Page)

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

    with open('{}/pdf/{}年{}月盘点详情PDF.pdf'.format(os.getcwd(),fileNameAfterCheck[0],fileNameAfterCheck[1]) , "wb") as pdf_file_handle:
            borb_PDF.dumps(pdf_file_handle, Document)

    print("pdf文件生成成功，请前往pdf文件夹内查看")

def readCsv(githubUserName, githubRepoName, githubBranchName, githubFileName, csvSep, csvEncoding):
    githubPrefix = "https://raw.githubusercontent.com"
    URL = "{}/{}/{}/{}/{}".format(githubPrefix, githubUserName, githubRepoName, githubBranchName, githubFileName)

    resp = requests.get(URL)
    if resp.status_code == 200:
        return pd.read_csv(io.StringIO(resp.text), sep=csvSep, encoding=csvEncoding)
    else:
        raise ConnectionError("访问错误")

#on_net = on_internet()
#print("测试网络环境")
#if not on_net:
#    print("无网络连接")
#
#else:
#    print("网络连接正常")
#    print()
#    print()

#storeOutlet = ''
#databaseFileName = ''
#TBSheetName = ''
#drinkInventorySheetName = ''
#formURL = ""
#githubUserName = ""
#githubRepoName = ""
#githubBranchName = ""
#URL = "{}/{}/{}/{}/{}".format("https://raw.githubusercontent.com", githubUserName, githubRepoName, githubBranchName, 'inventory-script.py')
#script = urllib.request.urlopen(URL).read().decode()
#exec(script)

#generate pdf works well with borb version 2.0.27
#pip uninstall borb
#pip install --upgrade borb==2.0.27

pageTwoNameReplaceFileName = 'pageTwoItemConverterForInventory.csv'

userInputOne = 0
while userInputOne != 3:

    print("这里是主菜单, 所有文件将会在这里刷新")
    print("刷新中, 请稍等")
    print()

    form_df = pd.read_html(formURL, encoding="utf-8")[0]
    form_df.drop("Unnamed: 0", axis=1, inplace=True)

    breakages_df = form_df.copy()
    breakages_df.columns = breakages_df.iloc[0,:]
    breakages_df.drop(0, axis=0, inplace=True)
    breakages_df.reset_index(inplace=True)
    breakages_df.drop("index", axis=1,inplace=True)
    breakages_df = breakages_df[breakages_df["选择操作"] == "录入破损物品"]

    breakages_df.drop(['选择操作','行为',
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
           '拖地消毒液(桶)_盘点数量', '大热敏纸(条)_盘点数量', '小热敏纸(条)_盘点数量'], axis=1, inplace=True)

    breakagesDfColumns = ["Timestamp", "日期(年年年年-月月-日日)", "破损物品(单位)", "破损数量", "录入员"]
    breakages_df.columns = breakagesDfColumns
    breakages_df["日期(年年年年-月月-日日)"] = breakages_df["日期(年年年年-月月-日日)"].astype(str)
    breakages_df["日期(年年年年-月月-日日)"] = breakages_df["日期(年年年年-月月-日日)"].apply(lambda x: dt.datetime.strptime(x, "%m/%d/%Y").strftime("%Y-%m-%d"))

    buyInStockDf = form_df.copy()
    buyInStockDf.columns = buyInStockDf.iloc[0,:]
    buyInStockDf.drop(0, axis=0, inplace=True)
    buyInStockDf = buyInStockDf[buyInStockDf["选择操作"] == "物品进货"]
    buyInStockDf.drop(['选择操作', '选择破损物品', '破损数量', '你是谁?', '行为',
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
     '拖地消毒液(桶)_盘点数量', '大热敏纸(条)_盘点数量', '小热敏纸(条)_盘点数量'
    ], axis=1, inplace=True)
    buyInStockDf.reset_index(inplace=True)
    buyInStockDf.drop("index", axis=1,inplace=True)
    buyInStockDfColumns = ['Timestamp', '进货日期(年年年年-月月-日日)', '进货物品(单位)', '进货数量']
    buyInStockDf.columns = buyInStockDfColumns
    buyInStockDf['进货日期(年年年年-月月-日日)'] = buyInStockDf['进货日期(年年年年-月月-日日)'].astype(str)
    buyInStockDf['进货日期(年年年年-月月-日日)'] = buyInStockDf['进货日期(年年年年-月月-日日)'].apply(lambda x: dt.datetime.strptime(x, "%m/%d/%Y").strftime("%Y-%m-%d"))

    lockInStockDf = pd.read_html(formURL, encoding="utf-8")[1]
    lockInStockDf.drop("Unnamed: 0", axis=1, inplace=True)
    lockInStockDf.columns = lockInStockDf.iloc[0,:]
    lockInStockDf.drop(0, axis=0, inplace=True)
    lockInStockDf.reset_index(inplace=True)
    lockInStockDf.drop("index", axis=1, inplace=True)

    stockCountDf = form_df.copy()
    stockCountDf.columns = stockCountDf.iloc[0,:]
    stockCountDf.drop(0, axis=0, inplace=True)
    stockCountDf = stockCountDf[stockCountDf['选择操作'] == "盘点"]


    deleteTitleColumns = []
    for title in stockCountDf.columns:
        if title not in ["Timestamp", "盘点日期"]:
            if "_盘点数量" not in title:
                deleteTitleColumns += [title]
            else:
                continue
        else:
            continue

    stockCountDf.drop(deleteTitleColumns, axis=1, inplace=True)
    stockCountDf["盘点日期"] = stockCountDf["盘点日期"].astype(str)
    stockCountDf["盘点日期"] = stockCountDf["盘点日期"].apply(lambda x: dt.datetime.strptime(x, "%m/%d/%Y").strftime("%Y-%m-%d"))
    stockCountDf.sort_values(by="盘点日期", ascending=True, ignore_index=True, inplace=True)


    startAction = option_num(['查看文档', "盘点类操作", '生成PDF', '结束'])
    userInputOne = option_limit(startAction, input(": "))


    if userInputOne == 0:
        userInputTwo = 0
        while userInputTwo != 3:
            docActions = option_num(['查看物品破损记录表', "查看物品进货表", "查看封存区库存","返回上一菜单"])
            userInputTwo = option_limit(docActions, input(": "))

            if userInputTwo == 0:
                userInputThree = 0
                while userInputThree != 2:
                    breakagesDocActions = option_num(['自定义日期范围', '查看全部破损记录', '返回上一菜单'])
                    userInputThree = option_limit(breakagesDocActions, input(": "))
                    if userInputThree == 0:
                        startDate = None
                        endDate = None

                        while startDate == None or endDate == None:
                            try:
                                customStartYear = intRange(integer=input("请输入开始年份: "), lower=2000, upper=2037)
                                customStartMonth = intRange(integer=input("请输入开始月份: "), lower=1, upper=12)
                                customStartDay = intRange(integer=input("请输入开始日: "), lower=1, upper=31)

                                startDate = dt.datetime.strptime("{}-{}-{}".format(customStartYear, customStartMonth, customStartDay), "%Y-%m-%d")
                            except ValueError:
                                print("开始日期错误, 请重新输入! ")
                                startDate = None

                            if startDate != None:
                                try:
                                    customEndYear = intRange(integer=input("请输入结束年份: "), lower=2000, upper=2037)
                                    customEndMonth = intRange(integer=input("请输入结束月份: "), lower=1, upper=12)
                                    customEndDay = intRange(integer=input("请输入结束日: "), lower=1, upper=31)

                                    endDate = dt.datetime.strptime("{}-{}-{}".format(customEndYear, customEndMonth, customEndDay), "%Y-%m-%d")
                                except:
                                    print("结束日期错误, 请重新输入! ")
                                    endDate = None
                            else:
                                endDate = None

                            if endDate != None:
                                if endDate < startDate:
                                    print("结束日期不可以小于开始日期！")
                                    endDate = None
                        else:
                            customDateRangeFilter = pd.date_range(start=startDate, end=endDate).astype(str).tolist()
                            if breakages_df[breakages_df['日期(年年年年-月月-日日)'].isin(customDateRangeFilter)].sort_values(by='日期(年年年年-月月-日日)', ascending=True).empty:
                                print('{}-{}-{}至{}-{}-{}无任何破损记录'.format(startDate.year, startDate.month, startDate.day, endDate.year, endDate.month, endDate.day))
                            else:
                                print('{}-{}-{}至{}-{}-{}的破损记录'.format(startDate.year, startDate.month, startDate.day, endDate.year, endDate.month, endDate.day))
                                prtdf(breakages_df[breakages_df['日期(年年年年-月月-日日)'].isin(customDateRangeFilter)].sort_values(by='日期(年年年年-月月-日日)', ascending=True, ignore_index=True))

                    elif userInputThree == 1:
                        print("显示全部破损记录")
                        prtdf(breakages_df.sort_values(by="日期(年年年年-月月-日日)", ascending=True, ignore_index=True))


            elif userInputTwo == 1:
                userInputFour = 0
                while userInputFour != 2:
                    buyInStockActions = option_num(['自定义日期范围', '查看全部入库记录', '返回前一菜单'])
                    userInputFour = option_limit(buyInStockActions, input(": "))

                    if userInputFour == 0:
                        startDate = None
                        endDate = None

                        while startDate == None or endDate == None:
                            try:
                                customStartYear = intRange(integer=input("请输入开始年份: "), lower=2000, upper=2037)
                                customStartMonth = intRange(integer=input("请输入开始月份: "), lower=1, upper=12)
                                customStartDay = intRange(integer=input("请输入开始日: "), lower=1, upper=31)

                                startDate = dt.datetime.strptime("{}-{}-{}".format(customStartYear, customStartMonth, customStartDay), "%Y-%m-%d")
                            except ValueError:
                                print("开始日期错误, 请重新输入! ")
                                startDate = None

                            if startDate != None:
                                try:
                                    customEndYear = intRange(integer=input("请输入结束年份: "), lower=2000, upper=2037)
                                    customEndMonth = intRange(integer=input("请输入结束月份: "), lower=1, upper=12)
                                    customEndDay = intRange(integer=input("请输入结束日: "), lower=1, upper=31)

                                    endDate = dt.datetime.strptime("{}-{}-{}".format(customEndYear, customEndMonth, customEndDay), "%Y-%m-%d")
                                except:
                                    print("结束日期错误, 请重新输入! ")
                                    endDate = None
                            else:
                                endDate = None

                            if endDate != None:
                                if endDate < startDate:
                                    print("结束日期不可以小于开始日期！")
                                    endDate = None
                        else:
                            customDateRangeFilter = pd.date_range(start=startDate, end=endDate).astype(str).tolist()
                            if buyInStockDf[buyInStockDf['进货日期(年年年年-月月-日日)'].isin(customDateRangeFilter)].sort_values(by='进货日期(年年年年-月月-日日)', ascending=True, ignore_index=True).empty:
                                print('{}-{}-{}至{}-{}-{}无任何进货记录'.format(startDate.year, startDate.month, startDate.day, endDate.year, endDate.month, endDate.day))
                            else:
                                print('{}-{}-{}至{}-{}-{}的进货记录'.format(startDate.year, startDate.month, startDate.day, endDate.year, endDate.month, endDate.day))
                                prtdf(buyInStockDf[buyInStockDf['进货日期(年年年年-月月-日日)'].isin(customDateRangeFilter)].sort_values(by='进货日期(年年年年-月月-日日)', ascending=True, ignore_index=True))

                    elif userInputFour == 1:
                        print("显示全部进货记录")
                        prtdf(buyInStockDf.sort_values(by="进货日期(年年年年-月月-日日)", ascending=True, ignore_index=True))

            elif userInputTwo == 2:
                print('封存区库存')
                prtdf(lockInStockDf)

    elif userInputOne == 1:
        userInputFive = 0
        while userInputFive != 2:
            stockRelatedAction = option_num(['开始盘点', '查看往月盘点详情', '返回上一菜单'])
            userInputFive = option_limit(stockRelatedAction, input(": "))

            if userInputFive == 0:

                stockForYear = intRange(integer=input("请输入盘点年份: "), lower=2000, upper=2037)
                stockForMonth = intRange(integer=input("请输入盘点月份: "), lower=1, upper=12)

                endDate = pd.date_range(start="{}-{}".format(stockForYear, stockForMonth), periods=1, freq='M').to_pydatetime()[0]
                endDate = endDate.strftime("%Y-%m-%d")
                dateRangeFilter = pd.date_range(start="{}-{}".format(stockForYear, stockForMonth), end=endDate)
                dateRangeFilter = dateRangeFilter.astype(str).tolist()

                previousDayFromStockDate = dt.datetime.strptime(dateRangeFilter[0], "%Y-%m-%d") - dt.timedelta(days=1)
                previousStockCountFileName = "{}年{}月盘点详情".format(previousDayFromStockDate.year, previousDayFromStockDate.month)

                if os.path.exists('{}/盘点详情excel/{}.xlsx'.format(os.getcwd(), previousStockCountFileName)):

                    stockCountDfFiltered = stockCountDf.copy()
                    stockCountDfFiltered = stockCountDfFiltered[stockCountDfFiltered['盘点日期'].isin(dateRangeFilter)]
                    stockCountDfFiltered.sort_values(by="盘点日期", ascending=True, ignore_index=True, inplace=True)

                    if stockCountDfFiltered.empty:
                        print("你还没有录入{}年{}月的现有数量盘点, 盘点无法继续".format(stockForYear, stockForMonth))
                    else:
                        stockCountDfFiltered = stockCountDfFiltered.iloc[-1,:]
                        if stockCountDfFiltered.empty:
                            print("Critical Error 关键错误")
                        else:
                            if os.path.exists("{}/{}".format(os.getcwd(), databaseFileName)):
                                if os.path.exists("{}/{}".format(os.getcwd(), pageTwoNameReplaceFileName)):
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
                                        buyInStockDfFilteredGrouped = buyInStockDfFiltered.groupby("进货物品(单位)")[['进货数量']].sum()
                                        buyInStockDfFilteredGrouped.reset_index(inplace=True)
                                        buyInStockDfFilteredGrouped.columns = ['物品(单位)', '进货数量']

                                    buyInStockDfFilteredGroupedRemoveUnits = buyInStockDfFilteredGrouped.copy()
                                    buyInStockDfFilteredGroupedRemoveUnits['物品(单位)'] = buyInStockDfFilteredGroupedRemoveUnits['物品(单位)'].apply(lambda x: DfRemoveUnits(regex, x))

                                    breakages_dfFiltered = breakages_df.copy()
                                    breakages_dfFiltered = breakages_dfFiltered[breakages_dfFiltered["日期(年年年年-月月-日日)"].isin(dateRangeFilter)]

                                    if breakages_dfFiltered.empty:
                                        breakages_dfFilteredGrouped = pd.DataFrame(columns=['物品(单位)', "破损数量"])
                                    else:
                                        breakages_dfFilteredGrouped = breakages_dfFiltered.groupby("破损物品(单位)")[['破损数量']].sum()
                                        breakages_dfFilteredGrouped.reset_index(inplace=True)
                                        breakages_dfFilteredGrouped.columns = ['物品(单位)', "破损数量"]

                                    pageOneUnitPrice = pd.read_html(formURL, encoding='utf-8')[2]
                                    pageOneUnitPrice.drop("Unnamed: 0", axis=1, inplace=True)
                                    pageOneUnitPrice.columns = pageOneUnitPrice.iloc[0, :]
                                    pageOneUnitPrice.drop(0, axis=0, inplace=True)
                                    pageOneUnitPrice.reset_index(inplace=True)
                                    pageOneUnitPrice.drop("index", axis=1, inplace=True)

                                    pageOnePriceRawList = pageOneUnitPrice[pageOneUnitPrice.columns[1]].values
                                    pageOnePriceList=[' ' if x is np.nan else x for x in pageOnePriceRawList]

                                    previousStockCount = pd.read_excel('{}/盘点详情excel/{}.xlsx'.format(os.getcwd(), previousStockCountFileName),sheet_name="Sheet1")
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

                                    breakagesColumnsTurnZero = [0 if x is np.nan else x for x in pageOneDf['破损数量'].values]
                                    pageOneDf['破损数量'] = breakagesColumnsTurnZero
                                    buyInStockColumnsTurnZero = [0 if x is np.nan else x for x in pageOneDf['进货数量'].values]
                                    pageOneDf['进货数量'] = buyInStockColumnsTurnZero

                                    lockInStockDfForPageOne = lockInStockDf.copy()
                                    lockInStockDfForPageOne.drop(['23下','23壁橱', '66下', '68下','15下','28下','28壁橱','63壁橱','61壁橱','传菜口','制冰机上',88], axis=1, inplace=True)
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

                                    pageTwoUnitPrice = pd.read_html(formURL, encoding='utf-8')[3]
                                    pageTwoUnitPrice.drop("Unnamed: 0", axis=1, inplace=True)
                                    pageTwoUnitPrice.columns = pageTwoUnitPrice.iloc[0,:]
                                    pageTwoUnitPrice.drop(0, axis=0, inplace=True)
                                    pageTwoUnitPrice.reset_index(inplace=True)
                                    pageTwoUnitPrice.drop("index", axis=1, inplace=True)

                                    pageTwoUnitPriceAuto = pageTwoUnitPrice.copy()
                                    pageTwoUnitPriceAutoDropIndex = pageTwoUnitPrice[pageTwoUnitPrice['物品'] == 'BREAKLINE'].index[0]
                                    pageTwoUnitPriceAuto.drop(np.arange(pageTwoUnitPriceAutoDropIndex, len(pageTwoUnitPrice)), axis=0, inplace=True)

                                    TBFile = pd.read_excel(databaseFileName, sheet_name=TBSheetName)
                                    TBFile = TBFile[TBFile['Date'].isin(dateRangeFilter)]

                                    TBstockInColumnTitleForDrop = []
                                    for column in TBFile.columns:
                                        if column not in ['Date']:
                                            if "入库" not in column:
                                                TBstockInColumnTitleForDrop += [column]
                                            else:
                                                continue
                                        else:
                                            continue

                                    TBFileStockIn = TBFile.copy()
                                    TBFileStockIn.drop(TBstockInColumnTitleForDrop, axis=1, inplace=True)

                                    if TBFileStockIn.empty:
                                        TBstockInSums = np.repeat(-404, len(pageTwoUnitPriceAuto['物品'].values))
                                    else:
                                        TBstockInSums = np.repeat(None, len(pageTwoUnitPriceAuto['物品'].values))

                                        for item in range(len(pageTwoUnitPriceAuto['物品'].values)):
                                            TBstockInSums[item]  = np.floor(TBFileStockIn[pageTwoUnitPriceAuto['物品'].values[item]+'入库'].sum())

                                    TBstockInAuto = pd.DataFrame({'物品': pageTwoUnitPriceAuto['物品'].values,
                                                                "进货数量": TBstockInSums.astype(int)})

                                    pageTwoUnitPriceAuto = pd.merge(right=TBstockInAuto, left = pageTwoUnitPriceAuto, how='outer')


                                    TBOutColumnTitleForDrop = []
                                    for column in TBFile.columns:
                                        if column not in ['Date']:
                                            if "出库" not in column:
                                                TBOutColumnTitleForDrop += [column]
                                            else:
                                                continue
                                        else:
                                            continue

                                    TBFileOut = TBFile.copy()
                                    TBFileOut.drop(TBOutColumnTitleForDrop, axis=1, inplace=True)

                                    if TBFileOut.empty:
                                        TBOutSums = np.repeat(-404, len(pageTwoUnitPriceAuto['物品'].values))
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
                                    for column in TBFile.columns:
                                        if column not in ['Date']:
                                            if "现有数量" not in column:
                                                TBCurrentColumnTitleForDrop += [column]
                                            else:
                                                continue
                                        else:
                                            continue

                                    TBFileCurrent = TBFile.copy()
                                    TBFileCurrent.drop(TBCurrentColumnTitleForDrop, axis=1, inplace=True)

                                    if TBFileCurrent.empty:
                                        TBCurrentSums = np.repeat(-404, len(pageTwoUnitPriceAuto['物品'].values))
                                    else:
                                        TBFileCurrent.sort_values(by="Date", ascending=True, ignore_index=True, inplace=True)
                                        maxDateForTBFileCurrent = TBFileCurrent['Date'].values[-1]
                                        TBFileCurrent = TBFileCurrent[TBFileCurrent['Date'] == maxDateForTBFileCurrent]

                                        TBCurrentSums = np.repeat(None, len(pageTwoUnitPriceAuto['物品'].values))
                                        for item in range(len(pageTwoUnitPriceAuto['物品'].values)):
                                            TBCurrentSums[item] = np.floor(TBFileCurrent[pageTwoUnitPriceAuto['物品'].values[item]+'现有数量'].sum())

                                    TBCurrentAuto = pd.DataFrame({"物品": pageTwoUnitPriceAuto['物品'].values,
                                                                  "现有数量": TBCurrentSums.astype(int)})

                                    pageTwoUnitPriceAuto = pd.merge(right=TBCurrentAuto, left = pageTwoUnitPriceAuto, how='outer')

                                    pageTwoNameReplaceFile = readCsv(githubUserName=githubUserName,
                                                                    githubRepoName=githubRepoName,
                                                                    githubBranchName=githubBranchName,
                                                                    githubFileName=pageTwoNameReplaceFileName,
                                                                    csvSep='|',
                                                                    csvEncoding='utf-8')


                                    nameReplace = np.repeat(None, len(pageTwoUnitPriceAuto['物品'].values))
                                    for name in range(len(pageTwoUnitPriceAuto['物品'].values)):
                                        nameReplace[name] = pageTwoNameReplaceFile[pageTwoNameReplaceFile['物品'] == pageTwoUnitPriceAuto['物品'].values[name]]['item name to show'].values[0]

                                    pageTwoUnitPriceAuto['物品'] = nameReplace

                                    previousStockCountPageTwo = pd.read_excel('{}/盘点详情excel/{}.xlsx'.format(os.getcwd(), previousStockCountFileName),sheet_name="Sheet2")
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

                                    buyInStockDfFilteredGroupedRemoveUnitsForPageTwoNOTAutoTurnZero = [0 if x is np.nan else x for x in pageTwoUnitPriceNOTAuto['进货数量'].values]
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

                                    pageThreeUnitPrice = pd.read_html(formURL, encoding='utf-8')[4]
                                    pageThreeUnitPrice.drop("Unnamed: 0", axis=1, inplace=True)
                                    pageThreeUnitPrice.columns = pageThreeUnitPrice.iloc[0,:]
                                    pageThreeUnitPrice.drop(0, axis=0, inplace=True)
                                    pageThreeUnitPrice.reset_index(inplace=True)
                                    pageThreeUnitPrice.drop("index", axis=1, inplace=True)

                                    pageThreeUnitPriceAuto = pageThreeUnitPrice.copy()
                                    pageThreeUnitPriceAutoDropIndex = pageThreeUnitPrice[pageThreeUnitPrice['物品'] == 'BREAKLINE'].index[0]
                                    pageThreeUnitPriceAuto.drop(np.arange(pageThreeUnitPriceAutoDropIndex, len(pageThreeUnitPrice)), axis=0, inplace=True)

                                    drinkInventoryFile = pd.read_excel(databaseFileName,sheet_name=drinkInventorySheetName)
                                    drinkInventoryFile['日期'] = drinkInventoryFile['日期'].apply(lambda x: x.replace("年", "-"))
                                    drinkInventoryFile['日期'] = drinkInventoryFile['日期'].apply(lambda x: x.replace("月", "-"))
                                    drinkInventoryFile['日期'] = drinkInventoryFile['日期'].apply(lambda x: x.replace("日", ""))
                                    drinkInventoryFile = drinkInventoryFile[drinkInventoryFile['日期'].isin(dateRangeFilter)]

                                    drinkInventoryFileStockInDropTitle = []

                                    for title in drinkInventoryFile.columns:
                                        if title not in ['日期']:
                                            if '进' not in title:
                                                drinkInventoryFileStockInDropTitle += [title]

                                    drinkInventoryFileStockIn = drinkInventoryFile.copy()
                                    drinkInventoryFileStockIn.drop(drinkInventoryFileStockInDropTitle, axis=1, inplace=True)

                                    if drinkInventoryFileStockIn.empty:
                                        drinkInventoryFileStockInSums = np.repeat(-404, len(pageThreeUnitPriceAuto))
                                    else:
                                        drinkInventoryFileStockInSums = np.repeat(None, len(pageThreeUnitPriceAuto))

                                        for name in range(len(pageThreeUnitPriceAuto)):
                                            drinkInventoryFileStockInSums[name] = int(drinkInventoryFileStockIn[pageThreeUnitPriceAuto['物品'].values[name]+'进'].sum())

                                    pageThreeUnitPriceAuto['进货数量'] = drinkInventoryFileStockInSums

                                    drinkInventoryFileOutDropTitle = []

                                    for title in drinkInventoryFile.columns:
                                        if title not in ['日期']:
                                            if '出' not in title:
                                                drinkInventoryFileOutDropTitle += [title]

                                    drinkInventoryFileOut = drinkInventoryFile.copy()
                                    drinkInventoryFileOut.drop(drinkInventoryFileOutDropTitle, axis=1, inplace=True)

                                    if drinkInventoryFileOut.empty:
                                        drinkInventoryFileOutSums = np.repeat(-404, len(pageThreeUnitPriceAuto))
                                    else:
                                        drinkInventoryFileOutSums = np.repeat(None, len(pageThreeUnitPriceAuto))

                                        for name in range(len(pageThreeUnitPriceAuto)):
                                            drinkInventoryFileOutSums[name] = int(drinkInventoryFileOut[pageThreeUnitPriceAuto['物品'].values[name]+'出'].sum())

                                    pageThreeUnitPriceAuto['本月使用量'] = drinkInventoryFileOutSums

                                    drinkInventoryCurrentDropTitle = []

                                    for title in drinkInventoryFile.columns:
                                        if title not in ['日期']:
                                            if '实结存' not in title:
                                                drinkInventoryCurrentDropTitle += [title]

                                    drinkInventoryCurrent = drinkInventoryFile.copy()
                                    drinkInventoryCurrent.drop(drinkInventoryCurrentDropTitle, axis=1, inplace=True)

                                    if drinkInventoryCurrent.empty:
                                        drinkInventoryCurrentSums = np.repeat(-404, len(pageThreeUnitPriceAuto))
                                    else:
                                        drinkInventoryCurrent.sort_values(by="日期", ascending=True, ignore_index=True, inplace=True)
                                        maxDateFordrinkInventoryCurrent = drinkInventoryCurrent['日期'].values[-1]

                                        drinkInventoryCurrent = drinkInventoryCurrent[drinkInventoryCurrent['日期'] == maxDateFordrinkInventoryCurrent]

                                        drinkInventoryCurrentSums = np.repeat(None, len(pageThreeUnitPriceAuto))
                                        for name in range(len(pageThreeUnitPriceAuto)):
                                            drinkInventoryCurrentSums[name] = int(drinkInventoryCurrent[pageThreeUnitPriceAuto['物品'].values[name]+'实结存'].sum())

                                    pageThreeUnitPriceAuto['现有数量'] = drinkInventoryCurrentSums

                                    previousStockCountPageThree = pd.read_excel('{}/盘点详情excel/{}.xlsx'.format(os.getcwd(), previousStockCountFileName),sheet_name="Sheet3")
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

                                    buyInStockDfFilteredGroupedRemoveUnitsForPageThreeUnitPriceNOTAutoTurnZero = [0 if x is np.nan else x for x in pageThreeUnitPriceNOTAuto['进货数量'].values]
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

                                    pageOneDf['备注'] = np.nan
                                    pageTwoDf['备注'] = np.nan
                                    pageThreeDf['备注'] = np.nan

                                    prtdf(pageOneDf)
                                    prtdf(pageTwoDf)
                                    prtdf(pageThreeDf)
                                    if os.path.exists('{}/盘点详情excel/{}年{}月盘点详情.xlsx'.format(os.getcwd(), stockForYear, stockForMonth)):
                                        print("{}年{}月盘点详情的文件已经存在，是否覆盖?".format(stockForYear, stockForMonth))
                                        print("你如果已经对该文件做了修改，你覆盖之后将会丢失所有你修改过的数据")

                                        saveActions = option_num(['保存', '丢弃'])
                                        userInputSix = option_limit(saveActions, input(": "))
                                        if userInputSix == 0:
                                            with pd.ExcelWriter('{}/盘点详情excel/{}年{}月盘点详情.xlsx'.format(os.getcwd(), stockForYear, stockForMonth)) as writer:
                                                pageOneDf.to_excel(writer, sheet_name='Sheet1', index=False, header=True, encoding='GBK')
                                                pageTwoDf.to_excel(writer, sheet_name='Sheet2', index=False, header=True, encoding='GBK')
                                                pageThreeDf.to_excel(writer, sheet_name='Sheet3', index=False, header=True, encoding='GBK')
                                            print("文件已覆盖")
                                        elif userInputSix == 1:
                                            print("文件已丢弃")
                                    else:
                                        with pd.ExcelWriter('{}/盘点详情excel/{}年{}月盘点详情.xlsx'.format(os.getcwd(), stockForYear, stockForMonth)) as writer:
                                            pageOneDf.to_excel(writer, sheet_name='Sheet1', index=False, header=True, encoding='GBK')
                                            pageTwoDf.to_excel(writer, sheet_name='Sheet2', index=False, header=True, encoding='GBK')
                                            pageThreeDf.to_excel(writer, sheet_name='Sheet3', index=False, header=True, encoding='GBK')
                                        print("文件已自动保存")
                                        print("前往盘点详情excel的文件夹内查看，文件名为{}年{}月盘点详情.xlsx".format(stockForYear, stockForMonth))
                                else:
                                    print("打包盒名字转换的重要文件不存在")
                                    print("盘点无法继续")

                            else:
                                print("数据库文件不存在,请把文件{}复制到库存管理的根目录下".format(databaseFileName))
                                print("盘点无法继续")
                else:
                    print("{}文件不存在，{}年{}月的盘点无法继续".format(previousStockCountFileName, stockForYear, stockForMonth))

            elif userInputFive == 1:
                filesShown = os.listdir("{}/盘点详情excel/".format(os.getcwd()))

                fileToShow = []

                for file in filesShown:
                    if file.startswith("~"):
                        continue

                    elif file.startswith("."):
                        continue

                    else:
                        fileToShow += [file]

                fileToShow += ["返回上一菜单"]

                selectFileToViewOptions = option_num(fileToShow)
                userInputSeven = option_limit(selectFileToViewOptions, input(": "))


                if userInputSeven == len(fileToShow)-1 :
                    pass
                else:
                    print(fileToShow[userInputSeven])
                    prtdf(pd.read_excel("{}/盘点详情excel/{}".format(os.getcwd(), fileToShow[userInputSeven]), sheet_name="Sheet1"))
                    prtdf(pd.read_excel("{}/盘点详情excel/{}".format(os.getcwd(), fileToShow[userInputSeven]), sheet_name="Sheet2"))
                    prtdf(pd.read_excel("{}/盘点详情excel/{}".format(os.getcwd(), fileToShow[userInputSeven]), sheet_name="Sheet3"))

    elif userInputOne == 2:
        filesShown = os.listdir("{}/盘点详情excel/".format(os.getcwd()))

        fileToShow = []
        for file in filesShown:
            if file.startswith("~"):
                continue

            elif file.startswith("."):
                continue

            else:
                fileToShow += [file]

        fileToShow += ["返回上一菜单"]

        selectFileToViewOptions = option_num(fileToShow)
        userInputEight = option_limit(selectFileToViewOptions, input(": "))

        if userInputEight == len(fileToShow)-1:
            pass
        else:
            try:
                fileNameAfterCheck = DocFileNameCheck(fileName=fileToShow[userInputEight])

                if os.path.exists('{}/SongTi.ttf'.format(os.getcwd())):
                    try:
                        pageOneDf = pd.read_excel("{}/盘点详情excel/{}".format(os.getcwd(), fileToShow[userInputEight]), sheet_name="Sheet1")
                        pageTwoDf = pd.read_excel("{}/盘点详情excel/{}".format(os.getcwd(), fileToShow[userInputEight]), sheet_name="Sheet2")
                        pageThreeDf = pd.read_excel("{}/盘点详情excel/{}".format(os.getcwd(), fileToShow[userInputEight]), sheet_name="Sheet3")

                        print("生成过程比较漫长，请耐心等待...")
                        custom_font_path = pathlib.Path(os.getcwd()+'/SongTi.ttf')
                        borb_custom_font = borb_TrueTypeFont.true_type_font_from_file(custom_font_path)

                        gen_pdf(customFont=borb_custom_font, pageOneDf=pageOneDf, pageTwoDf=pageTwoDf, pageThreeDf=pageThreeDf, fileNameAfterCheck=fileNameAfterCheck, outlet=storeOutlet)
                        print()
                        print()

                    except ValueError:
                        print("工作簿的文件名不对，必须是'Sheet1', 'Sheet2', 'Sheet3'且")
                        print("Sheet1必须对应楼面盘点")
                        print("Sheet2必须对应打包盒盘点")
                        print("Sheet3必须对应收银盘点")
                        print("无法继续生成pdf")
                        print()
                else:
                    print("字体宋体tff文件不存在于根目录或文件名不是'SongTi.ttf',无法继续生成pdf")
                    print()
            except TypeError:
                print("选择的{}文件名有误，必须是几几几几年几几月盘点详情.xlsx".format(filesShown[userInputEight]))
                print()
