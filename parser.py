import csv
import sys
import os
import time

import openpyxl
from openpyxl import Workbook
import pandas as pd

csv.field_size_limit(sys.maxsize)


def loadData():
    directory = os.fsencode(os.getcwd() + '/data')
    for subdir in os.listdir(directory):
        filename = os.path.join(directory, subdir).decode()
        if filename.endswith(".xlsx"):
            parseExcel(filename)


# aux function to compare 2 strings without spaces and casing
def bareCompare(str1, str2):
    bare1 = "".join(str1.split())
    bare2 = "".join(str2.split())

    return bare1.lower() == bare2.lower()


def parseExcel(filename):
    workbook = openpyxl.load_workbook(filename)
    ws = workbook.worksheets[0]
    for index, row in enumerate(ws.iter_rows(min_row=2)):
        print(row)

    workbook.close()


loadData()
