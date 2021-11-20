import csv
import sys
import os
import time
from copy import copy
from datetime import datetime
from functools import reduce

from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils import get_column_letter

import openpyxl
from openpyxl import Workbook
import pandas as pd

csv.field_size_limit(sys.maxsize)


def formatWS(ws):
    for rows in ws.iter_rows(min_row=1, max_row=1, min_col=1):
        for cell in rows:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="8EA9DB", end_color='8EA9DB', fill_type="solid")

    dims = {}
    for row in ws.rows:
        for cell in row:
            cell.alignment = Alignment(wrapText=True)
            alignment_obj = copy(cell.alignment)
            alignment_obj.horizontal = 'center'
            alignment_obj.vertical = 'center'
            cell.alignment = alignment_obj

            if cell.value:
                dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value))))
    for col, value in dims.items():
        ws.column_dimensions[col].width = value


def loadData():
    directory = os.fsencode(os.getcwd() + '/data')
    for subdir in os.listdir(directory):
        if not subdir.decode().startswith('~'):
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
    strList = cleanList(strList)
    return "\n".join(strList)


def cleanList(strList):
    return list(map(lambda x: x.strftime('%d/%m/%Y') if isinstance(x, datetime) else str(x), strList))


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
            if dataset[2] not in ssaNames:
                ssaNum += 1
            if dataset[4] not in programNames:
                # roles seems to be how many unique program entries
                roles += 1
            numVolunteer = dataset[9]
            ssaNames.append(dataset[2])
            centreNames.append(dataset[3])
            programNames.append(dataset[4])
            eventDates.append(dataset[10])
            eventTime.append(dataset[12])
            hoursPerSess.append(dataset[13])
            didVolunteerPlan.append(dataset[16])
            remarks.append(dataset[17])

        sessions = len(programNames)
        totalHoursVolunteer = reduce(lambda a, b: cleanNum(a) + cleanNum(b), hoursPerSess)
        entryData = [indGrp, number, grpName, status, numVolunteer, stackString(ssaNames), stackString(centreNames),
                     ssaNum, stackString(programNames), roles, stackString(eventDates), sessions,
                     stackString(eventTime),
                     stackString(hoursPerSess), totalHoursVolunteer, stackString(didVolunteerPlan),
                     stackString(remarks)]
        ws.append(entryData)
    formatWS(ws)
    return wb


def genProgramSheet(wb, dataDict):
    sheetName = '(G) Programme Details'
    header = ['S/N', 'Programme', 'Partners Involved', 'Programme Details (Date and Time)',
              'Number of Beneficiaries/Service Users', 'Number of Volunteers', '', '', 'Number of Volunteering Hours']

    subHeader = ['', '', '', '', '', 'Adhoc', 'Regular', 'Leader', '']

    ws = wb.create_sheet(sheetName)
    ws.title = sheetName
    ws.append(header)
    ws.append(subHeader)
    # header loop
    # there must be a smarter way to do this but... i don't know
    ws.merge_cells('A1:A2')
    ws.merge_cells('B1:B2')
    ws.merge_cells('C1:C2')
    ws.merge_cells('D1:D2')
    ws.merge_cells('A1:A2')
    ws.merge_cells('E1:E2')
    ws.merge_cells('F1:H1')
    ws.merge_cells('I1:I2')

    # proceed official data dissection here
    counter = 1
    for key in dataDict:
        dataArray = dataDict.get(key)
        progDates = []
        progTimes = []
        centres = []
        serviceUsers = 0
        adhocCount = 0
        regularCount = 0
        leaderCount = 0
        for i in range(len(dataArray)):
            dataset = dataArray[i]
            progName = dataset[4]
            partnerName = dataset[2] + " & " + dataset[8]
            if dataset[3] not in centres:
                centres.append(dataset[3])
                serviceUsers += dataset[15]
            progDates.append(dataset[10])
            progTimes.append(dataset[12])
            hoursPerSess = dataset[13]
            volunteernum = dataset[9]

        if len(progDates) <= 1:
            adhocCount += volunteernum
        else:
            regularCount += volunteernum

        totalHoursVolunteer = hoursPerSess * len(dataArray) * volunteernum
        progDates = cleanList(progDates)

        progDatetime = []
        for j in range(len(progDates)):
            progDate = progDates[j]
            progTime = progTimes[j]
            datetimeConcat = progDate + ', ' + progTime
            progDatetime.append(datetimeConcat)

        dataEntry = [counter, progName, partnerName, stackString(progDatetime), serviceUsers, adhocCount,
                     regularCount, leaderCount, totalHoursVolunteer]

        ws.append(dataEntry)
        counter += 1
    formatWS(ws)

    return wb


def parseExcel(filename):
    # I should be able to generate 4 additional sheets after being provided one sheet
    workbook = openpyxl.load_workbook(filename)
    outbook = Workbook()
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

    outbook = genVolunteerSheet(outbook, volunteerDict)
    outbook = genProgramSheet(outbook, programDict)
    del outbook['Sheet']

    outbook.save('output-data.xlsx')
    workbook.close()
    outbook.close()


loadData()
