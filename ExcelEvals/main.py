import pandas as pd
import numpy as np


# Script also uses openpyxl internal library for pandas, maybe xlrd (probably defunct)
def initialize(sheet, x):
    values = list()
    for index, row in sheet.iterrows():
        values.append(row[sheet.columns.values[x]])
    return values


def excel_mean(values):
    output = np.mean(values)
    print('Mean:', output)


def excel_median(values):
    output = np.median(values)
    print('Median:', output)


def excel_range(values):
    output = np.ptp(values)
    print('Range:', output)


def excel_std(values):
    output = np.std(values)
    print('Standard Deviation:', output)


def excel_variance(values):
    output = np.var(values)
    print('Variance:', output)


def base_stats(values):
    excel_mean(values)
    excel_median(values)
    excel_range(values)
    excel_variance(values)
    excel_std(values)


def analyze():
    path = input('Input Path to xlsx file: ')
    df = pd.read_excel(path)
    print(df)
    column = int(input('Please enter the column number with numerical data to be analyzed: '))
    values = initialize(df, column)
    base_stats(values)


hold = 1
while hold:
    answer = input('Type Y to analyze .xlsx file, type N to exit...')
    if answer.upper() == 'Y':
        analyze()
    elif answer.upper() == 'N':
        hold = 0
    else:
        print('\nNot a valid input')
        input()