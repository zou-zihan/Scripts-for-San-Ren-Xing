#! /usr/bin/env python3

import pandas as pd
import numpy as np
import datetime as dt
import math
import os

def parseGoogleHTMLSheet(df):

    df.drop("Unnamed: 0", axis=1, inplace=True)
    df.columns = df.iloc[0,:]
    df.drop(0, axis=0,inplace=True)

    col_1st = df.columns[0]
    df.drop(df[df[col_1st].isnull()].index,axis=0, inplace=True)

    df.reset_index(inplace=True)
    df.drop("index", axis=1, inplace=True)

    return df

def normal_round(n, decimals=0):
    expoN = n * 10 ** decimals
    if abs(expoN) - abs(math.floor(expoN)) < 0.5:
        return math.floor(expoN) / 10 ** decimals
    return math.ceil(expoN) / 10 ** decimals

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

def prtdf(df):
    with pd.option_context('display.max_rows', None,
                           'display.max_columns', None,
                           'display.width', 1000,
                           'display.precision', 2,
                           'display.colheader_justify', 'center'):
        return display(df)
    
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

#url = ""
#outlet=""
shift_database_filename = "{}_shift_database.xlsx".format(outlet)

if not os.path.exists(os.getcwd() + "/{}".format(shift_database_filename)):
    print("'{}'文件不存在".format(shift_database_filename))

else:
    userInputOne = 0
    while userInputOne != 2:
        print("这里是主菜单, 刷新中, 请稍等。")
        print()
        
        #shift_database = pd.DataFrame(columns=["FOR VLOOKUP","DATE", "DAY NAME", "ID", "NAME", "SHIFT", "OIL", "PH", "AL", "CCL", "REMARKS"])
        shift_database = pd.read_excel(shift_database_filename)
        shift_database["DATE"] = pd.to_datetime(shift_database["DATE"])
        shift_database["TIME LOG"] = pd.to_datetime(shift_database["TIME LOG"])
        shift_database["ID"] = shift_database["ID"].astype(int)
        shift_database["ID"] = shift_database["ID"].astype(str)

        df = pd.read_html(url, encoding="utf-8")
        
        shift_df = df[0]
        parseGoogleHTMLSheet(shift_df)
        shift_df["DATE"] = pd.to_datetime(shift_df["DATE"])
        shift_df["SHIFT"] = shift_df["SHIFT"].astype(str)
        shift_df["ID"] = shift_df["ID"].astype(int)
        shift_df["ID"] = shift_df["ID"].astype(str)

        employee_info = df[2]
        parseGoogleHTMLSheet(employee_info)
        employee_info["ID"] = employee_info["ID"].astype(int)
        employee_info["ID"] = employee_info["ID"].astype(str)
        employee_info["FIRST DAY DATE"] = pd.to_datetime(employee_info["FIRST DAY DATE"])
        employee_info["AL START DATE"] = pd.to_datetime(employee_info["AL START DATE"])
        employee_info["AL END DATE"] = pd.to_datetime(employee_info["AL END DATE"])

        ph_dates_df = df[3]
        parseGoogleHTMLSheet(ph_dates_df)
        ph_dates_df["PH DATE"] = pd.to_datetime(ph_dates_df["PH DATE"])

        leaves_manual_df = df[4]
        parseGoogleHTMLSheet(leaves_manual_df)
        leaves_manual_df["DATE"] = pd.to_datetime(leaves_manual_df["DATE"])
        leaves_manual_df["ID"] = leaves_manual_df["ID"].astype(int)
        leaves_manual_df["ID"] = leaves_manual_df["ID"].astype(str)
        
        print("刷新完成")
        print()
        
        startAction = option_num(["写入排班表数据库", "工资单时间分析", "结束"])
        userInputOne = option_limit(startAction, input(": "))
        
        if userInputOne == 0:
            print("开始写入...")

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
            shift_database.sort_values(by=["TIME LOG"], ascending=False, inplace=True)
            shift_database.reset_index(inplace=True)
            shift_database.drop("index", axis=1, inplace=True)
            
            dropIndex = shift_database[shift_database.duplicated(subset=["ID", "DATE"], keep="first")].index
            
            if len(dropIndex) > 0:
                shift_database.drop(dropIndex, axis=0, inplace=True)
                shift_database.sort_values(by=["TIME LOG"], ascending=False, inplace=True)
                shift_database.reset_index(inplace=True)
                shift_database.drop("index", axis=1, inplace=True)
            
            else:
                pass

            shift_database["DATE"] = shift_database["DATE"].apply(lambda x: x.strftime("%Y-%m-%d"))
            shift_database["IS PH"] = shift_database["IS PH"].apply(lambda x: str(x).strip().upper())
            shift_database["TIME LOG"] = shift_database["TIME LOG"].apply(lambda x: x.strftime("%Y-%m-%d %H:%M:%S"))
            
            shift_database.to_excel(shift_database_filename, index=False, header=True)
            
            print("写入完成。")
            
            prtdf(shift_database.head(20))
    
        elif userInputOne == 1:
            
            shift_database = pd.read_excel(shift_database_filename)
            shift_database["DATE"] = pd.to_datetime(shift_database["DATE"])
            shift_database["ID"] = shift_database["ID"].astype(int)
            shift_database["ID"] = shift_database["ID"].astype(str)
            
            userInputTwo = 0
            while userInputTwo != 3:
                payslip_action = option_num(["整月分析", "自定义日期范围分析", "全部记录分析", "返回上一菜单"])
                userInputTwo = option_limit(payslip_action, input(": "))
                
                if userInputTwo == 0:                
                    inputYear = intRange(integer=input("输入年份: "), lower=2000, upper=2037)
                    
                    inputMonth = intRange(integer=input("输入月份: "), lower=1, upper=12)
                    
                    start_date = dt.datetime(year=int(inputYear), month=int(inputMonth), day=1)
                    
                    end_date = dt.datetime(year=int(inputYear), month=int(inputMonth), day=month_last_day(inputYear, inputMonth))
                    
                    
                elif userInputTwo == 1:
                    start_date = None
                    end_date = None
                    
                    while start_date == None or end_date == None:
                        try:
                            start_year = intRange(integer=input("请输入开始年份: "), lower=2000, upper=2037)
                            start_month = intRange(integer=input("请输入开始月份: "), lower=2000, upper=2037)
                            start_day = intRange(integer=input("请输入开始日: "), lower=1, upper=31)
                            
                            start_date = dt.datetime(year=int(start_year), month=int(start_month), day = int(start_day))
                        
                        except ValueError:
                            print("开始日期错误, 请重新输入! ")
                            start_date = None
                        
                        if start_date != None:
                            try:
                                end_year = intRange(integer=input("请输入结束年份: "), lower=2000, upper=2037)
                                end_month = intRange(integer=input("请输入结束月份: "), lower=2000, upper=2037)
                                end_day = intRange(integer=input("请输入结束日: "), lower=1, upper=31)
                            
                                end_date = dt.datetime(year=int(end_year), month=int(end_month), day=int(end_day))
                                
                            except:
                                print("结束日期错误, 请重新输入! ")
                                end_date = None
                            
                        else:
                            end_date = None
                        
                        if end_date != None:
                            if end_date < start_date:
                                print("结束日期不可以小于开始日期！")
                                end_date = None
                
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
                                                                             outlet))


                    print()
                    for k,i in statement_dict.items():
                        print(k, ":", i)
                    
                    print()
                    print()
                
