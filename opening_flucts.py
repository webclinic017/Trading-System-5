##find win rates for different ranges of opening price fluct percentages
from ib.opt import ibConnection, message
from ib.ext.Contract import Contract
from ib.ext.Order import Order
from ib.ext.ExecutionFilter import ExecutionFilter
from ib.ext.CommissionReport import CommissionReport
from ib.ext.TickType import TickType as tt
from time import sleep
import json  #3 imports below are for google sheets api
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import sys
import time
import subprocess
import re

scope = ['https://spreadsheets.google.com/feeds']
credentials_file = "/Users/jakezimmerman/Documents/python3credentials.json"
credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
gc = gspread.authorize(credentials)
wks = gc.open("RCP90")
workbook1 = wks.get_worksheet(12)


win_cells = workbook1.range("O38:O620")
win_cells_list = []
win_cells_two = workbook1.range("P38:P620")
win_cells_two_list = []
loss_cells = workbook1.range("Q38:Q620")
loss_cells_list = []
loss_cells_two = workbook1.range("R38:R620")
loss_cells_two_list = []
percent_up_down = workbook1.range("C38:C620")
passed_params = workbook1.range("D38:D620")


#1 less than 1, 2 between 1 and 3.7, 3 between 3.8 and 6.5, 4 bewteen 6.6 and 10, 5 above 10
range_one_total = 0
range_one_wins = 0
range_two_total = 0
range_two_wins = 0
range_three_total = 0
range_three_wins = 0
range_four_total = 0
range_four_wins = 0
range_five_total = 0
range_five_wins = 0
params_wins = 0

percentageList = []
for cell in percent_up_down:
    percentageList.append(cell.value)
paramsList = []
for cell in passed_params:
    paramsList.append(cell.value)

def getColNum(cellInfo):
    col = ''
    string = str(cellInfo)
    m = re.search('R(.+?) ', string)
    if (m):
        col = m.group(1)
        return (col[-2:])


def countPercentageWins(cellRange):
    global range_one_total, range_one_wins, range_two_total, range_two_wins, range_three_total
    global range_four_wins, range_five_total, range_five_wins, range_three_wins, range_four_total
    global params_wins, fails_wins
    y = 0
    for cell in cellRange:
        #print("value: " + cell.value + " percentage: " + percentageList[y])
        if (cell.value == '' or percentageList[y] == ''):
            y+=1
            continue
        elif(cell.value != '' and abs(float(percentageList[y])) < 1):
            if (getColNum(cell) == '15' or getColNum(cell) == '16'):
                range_one_total+=1
                range_one_wins+=1
                y+=1
            else:
                range_one_total+=1
                y+=1
        elif(cell.value != '' and (abs(float(percentageList[y])) >= 1 and abs(float(percentageList[y])) <= 5)):
            if (getColNum(cell) == '15' or getColNum(cell) == '16'):
                range_two_total+=1
                range_two_wins+=1
                y+=1
            else:
                range_two_total+=1
                y+=1
        elif(cell.value != '' and (abs(float(percentageList[y])) > 5 and abs(float(percentageList[y])) <= 15)):
            if (getColNum(cell) == '15' or getColNum(cell) == '16'):
                if (paramsList[y] == 'Y'):
                    params_wins+=1
                range_three_total+=1
                range_three_wins+=1
                y+=1
            else:
                range_three_total+=1
                y+=1
        elif(cell.value != '' and (abs(float(percentageList[y])) > 15 and abs(float(percentageList[y])) <= 20)):
            if (getColNum(cell) == '15' or getColNum(cell) == '16'):
                if (paramsList[y] == 'Y'):
                    params_wins+=1
                range_four_total+=1
                range_four_wins+=1
                y+=1
            else:
                range_four_total+=1
                y+=1
        elif(cell.value != '' and abs(float(percentageList[y])) > 20):
            if (getColNum(cell) == '15' or getColNum(cell) == '16'):
                if (paramsList[y] == 'Y'):
                    params_wins+=1
                range_five_total+=1
                range_five_wins+=1
                y+=1
            else:
                range_five_total+=1
                y+=1
        else:
            y+=1
            pass



iterables = (win_cells, win_cells_two, loss_cells, loss_cells_two)
for iterable in iterables:
    countPercentageWins(iterable)
        
print("range 1: " + str(range_one_total) + " wins: " + str(range_one_wins) + " win rate: " + str(round((range_one_wins/range_one_total) * 100, 2)))
print("key: " + str(range_two_total) + " wins: " + str(range_two_wins) + " win rate: " + str(round((range_two_wins/range_two_total) * 100, 2)))
print("range 3: " + str(range_three_total) + " wins: " + str(range_three_wins) + " win rate: " + str(round((range_three_wins/range_three_total * 100), 2)))
print("range 4: " + str(range_four_total) + " wins: " + str(range_four_wins) + " win rate: " + str(round((range_four_wins/range_four_total) * 100, 2)))
print("range 5: " + str(range_five_total) + " wins: " + str(range_five_wins) + " win rate: " + str(round((range_five_wins/range_five_total) * 100, 2)))
total_params = 0
for cell in paramsList:
    if (cell == 'Y'):
        total_params+=1
#print("total params trades: %d wins: %d win rate: %.2f" % (total_params, params_wins, (params_wins/total_params) * 100))








