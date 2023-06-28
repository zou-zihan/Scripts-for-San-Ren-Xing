#! /usr/bin/env python3

import pandas as pd
import numpy as np
import datetime as dt
from dateutil.relativedelta import relativedelta
import math

def month_difference(start_date, end_date):
    delta = relativedelta(end_date, start_date)

    res_months = delta.months + (delta.years * 12)

    return res_months

def normal_round(n, decimals=0):
    expoN = n * 10 ** decimals
    if abs(expoN) - abs(math.floor(expoN)) < 0.5:
        return math.floor(expoN) / 10 ** decimals
    return math.ceil(expoN) / 10 ** decimals

def point_five_round(x):
    return normal_round(normal_round(x / 0.5) * 0.5, -int(np.floor(np.log10(0.5))))

def calculate_AL(start_date, end_date, first_day, position, unpaid_months, bonus_or_penalty):
    
    #Starting AL days differs by job position
    if position.upper().strip() == "CREW":
        base = 8
    elif position.upper().strip() == "MANAGER":
        base = 12
    
    start_date = dt.datetime.strptime(start_date, "%Y-%m-%d")
    first_day = dt.datetime.strptime(first_day, "%Y-%m-%d")
    
    end_date = dt.datetime.strptime(end_date, "%Y-%m-%d")
    end_date += relativedelta(days=1)
    
    month_of_commitment = month_difference(first_day, end_date)
    
    number_of_months = month_difference(start_date, end_date)
    
    unpaid_months = float(unpaid_months)
    
    number_of_months -= unpaid_months
    
    bonus_or_penalty = float(bonus_or_penalty)
    
    if month_difference(start_date, end_date) > 12:
        return None

    else:
        if month_of_commitment < 3:
            AL_claimables = 0 + bonus_or_penalty
            AL_claimables = point_five_round(AL_claimables)

            if AL_claimables < 0:
                AL_claimables = 0

            return AL_claimables

        elif month_of_commitment <= 12:       
            AL_claimables = number_of_months / 12 * base
            AL_claimables += bonus_or_penalty
            AL_claimables = point_five_round(AL_claimables)

            return AL_claimables

        else:
            AL_days = int(np.floor(month_of_commitment)) / 12 + base - 1

            AL_claimables = number_of_months / 12 * AL_days
            AL_claimables += bonus_or_penalty
            AL_claimables = point_five_round(AL_claimables)

            return AL_claimables

def prtdf(df):
    with pd.option_context('display.max_rows', None,
                           'display.max_columns', None,
                           'display.width', 1000,
                           'display.precision', 2,
                           'display.colheader_justify', 'center'):
        return display(df)

#url = ""

df = pd.read_html(url, encoding="utf-8")[0]
df.drop("Unnamed: 0", axis=1, inplace=True)
df.columns = df.iloc[0, :]
df.drop(0, axis=0, inplace=True)
df.reset_index(inplace=True)
df.drop("index", axis=1, inplace=True)

arr = np.zeros(len(df))

for index in range(len(df)):
    arr[index] = calculate_AL(start_date = df.iloc[index, 3],
                              end_date = df.iloc[index, 4],
                              first_day = df.iloc[index, 1],
                              position = df.iloc[index, 2],
                              unpaid_months = df.iloc[index, 5],
                              bonus_or_penalty = df.iloc[index, 6])
    
df["Annual Leave"] = arr

#form_url = ""

df1 = pd.read_html(form_url, encoding="utf-8")[0]
df1.drop("Unnamed: 0", axis=1, inplace=True)
df1.columns = df1.iloc[0,:].values.tolist()
df1.drop([0,1], axis=0, inplace=True)
df1.reset_index(inplace=True)
df1.drop("index", axis=1, inplace=True)

df1["Date of Birth 出生日期"] = pd.to_datetime(df1["Date of Birth 出生日期"], format="%m/%d/%Y")
df1.sort_values(by="Timestamp", ascending=False, inplace=True)

print("Annual Leave Calculator 年假计算器")
prtdf(df)

print()
print()
print()

print("Employee Information 员工资料")
prtdf(df1)
