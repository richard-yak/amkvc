import csv
import sys
import os
import time
from datetime import datetime
from functools import reduce

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
    return stripdown(str1) == stripdown(str2)


def stripdown(str1):
    return "".join(str1.split()).lower()


def cleanNum(num):
    try:
        return float(num)
    except ValueError:
        if '\n' in num:
            stringarr = num.split('\n')
            return float(stringarr[0])


def stackString(strList):
    strList = map(lambda x: x.strftime('%d/%m/%Y, %H:%M:%S') if isinstance(x, datetime) else str(x), strList)
    return "\n".join(strList)


def genVolunteerSheet(wb, dataDict):
    sheetName = '(D) Volunteers'
    header = ['Individual or Group', 'Individual or Group Number', 'Individual/Group name', 'Status',
              'No. of unique volunteers', 'Name of SSA', 'Centre Name', 'Number of SSAs', 'Programme / Activity name',
              'Number of Roles', 'Event Date', 'Number of Sessions', 'Event Time', 'Hours per session',
              'Total Hours Volunteered', 'Did the volunteer plan and coordinate the activity?',
              'Any additional Remarks']

    ws = wb.create_sheet(sheetName)
    ws.title = sheetName
    ws.append(header)

    for key in dataDict:
        rowData = []
        dataArray = dataDict.get(key)
        ssaNames = []
        centreNames = []
        programNames = []
        eventDates = []
        eventTime = []
        hoursPerSess = []
        didVolunteerPlan = []
        remarks = []
        roles = 0
        ssaNum = 0
        for i in range(len(dataArray)):
            dataset = dataArray[i]
            indGrp = dataset[5]
            number = dataset[7]
            grpName = dataset[8]
            status = ''
            numVolunteer = dataset[9]
            ssaNames.append(dataset[2])
            centreNames.append(dataset[3])
            programNames.append(dataset[4])

            if dataset[2] not in ssaNames:
                ssaNum += 1
            if dataset[4] not in programNames:
                # roles seems to be how many unique program entries
                roles += 1

            eventDates.append(dataset[10])
            eventTime.append(dataset[12])
            hoursPerSess.append(dataset[13])
            didVolunteerPlan.append(dataset[16])
            remarks.append(dataset[17])

        sessions = len(programNames)
        # THERE IS ERROR IN EXCEL, DOUBLE CHECK!!
        totalHoursVolunteer = reduce(lambda a, b: cleanNum(a) + cleanNum(b), hoursPerSess)
        entryData = [indGrp, number, grpName, status, numVolunteer, stackString(ssaNames), stackString(centreNames),
                     stackString(programNames), roles, stackString(eventDates), sessions, stackString(eventTime),
                     stackString(hoursPerSess), totalHoursVolunteer, didVolunteerPlan, stackString(remarks)]

        print (entryData)
    return wb


def genProgramSheet(wb, dataDict):
    return wb


def parseExcel(filename):
    # I should be able to generate 4 additional sheets after being provided one sheet
    workbook = openpyxl.load_workbook(filename)
    ws = workbook.worksheets[0]
    # index ref
    # [0] (Internal/External),[1] (AMK/SK/PG/Others),[2] (Name of SSA),[3] (Centre Name),
    # [4] (Programme / Activity name),[5] (Individual or Group),[6] (Type of Partner),
    # [7] (Individual or Group Number),[8] (Individual/Group name),[9] (No. of unique volunteers),
    # [10] (Event Date),[11] (No. of Sessions),[12] (Event time),[13] (Hours per session),
    # [14] (Total Number of Volunteering Hours),[15] (No. of Service users),
    # [16] (Did the volunteer plan and coordinate the activity?),[17] (Remarks),"
    dataArray = []
    volunteerDict = {}
    programDict = {}
    # preloaded
    for index, row in enumerate(ws.iter_rows(min_row=2)):
        dataArray = list(map(lambda x: x.value, row))
        # group by a few criteria, volunteer, partners, and program
        # print(dataArray)

        # make sure to stripdown all keys to prevent missed entries
        groupName = stripdown(dataArray[8])
        volunteerDict.setdefault(groupName, []).append(dataArray)
        programName = stripdown(dataArray[4])
        programDict.setdefault(programName, []).append(dataArray)

    workbook = genVolunteerSheet(workbook, volunteerDict)
    workbook = genProgramSheet(workbook, programDict)


    workbook.close()


loadData()
