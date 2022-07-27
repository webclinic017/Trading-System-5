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

credentials_file = "/Users/jakezimmerman/Documents/python3credentials.json"
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
gc = gspread.authorize(credentials)
wks = gc.open("RCP90")
which_data_set = "day_data_log"
columns = {"day_data_log" : ['L', 'M', 'N', 'O', 'I', 'X', "BUY", "SELL", "104", "326", 9, "12", "13"], "day_data_log" : ['O', 'P', 'Q', 'R', 'L', 'D', 'L', 'S', "77", "600", 12, "15", "16"]}
wins, wins_two, losses, losses_two, long_short, vix, long, short, beginning, end, sheet_num, win_col, win_two_col = 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12
data_log = wks.get_worksheet(columns[which_data_set][sheet_num])
#9 for current log, 12 for day data log

range_start = columns[which_data_set][beginning]
range_end = columns[which_data_set][end]
print("range start: ", range_start)

win_cells = data_log.range(columns[which_data_set][wins] + range_start + ':' + columns[which_data_set][wins] + range_end)
win_cells_list = []
win_cells_two  =data_log.range(columns[which_data_set][wins_two] + range_start + ':' + columns[which_data_set][wins_two] + range_end)
win_cells_two_list = []
loss_cells = data_log.range(columns[which_data_set][losses] + range_start + ':' + columns[which_data_set][losses] + range_end)
loss_cells_list = []
loss_cells_two = data_log.range(columns[which_data_set][losses_two] + range_start + ':' + columns[which_data_set][losses_two] + range_end)
loss_cells_two_list = []
long_short_cells = data_log.range(columns[which_data_set][long_short] + range_start + ':' + columns[which_data_set][long_short] + range_end)
long_short_list = []
vix_cells = data_log.range(columns[which_data_set][vix] + range_start + ':' + columns[which_data_set][vix] + range_end)
vix_list = []

def getColNum(cellInfo):
    col = ''
    string = str(cellInfo)
    m = re.search('R(.+?) ', string)
    if (m):
        col = m.group(1)
        return (col[-2:])

def countWins(iterable, pivot):
    global high_vix_long_wins, high_vix_long_losses, low_vix_long_wins, low_vix_long_losses
    global high_vix_short_wins, high_vix_short_losses, low_vix_short_wins, low_vix_short_losses
    y = 0
    for cell in iterable:
        if (cell.value == '' or vix_list[y] == '' or long_short_list[y] == ''):
            y+=1
            continue
        #win columns
        if (getColNum(cell) == columns[which_data_set][win_col] or getColNum(cell) == columns[which_data_set][win_two_col]):
            if (vix_list[y] >= pivot): #if vix was above 16
                if (long_short_list[y] == columns[which_data_set][long]):
                    high_vix_long_wins+=1
                elif (long_short_list[y] == columns[which_data_set][short]):
                    high_vix_short_wins+=1
                else:
                    pass
                y+=1
            else:
                if (long_short_list[y] == columns[which_data_set][long]):
                    low_vix_long_wins+=1
                elif (long_short_list[y] == columns[which_data_set][short]):
                    low_vix_short_wins+=1
                else:
                    pass
                y+=1
        #loss columns
        else:
            if (vix_list[y] >= pivot): 
                if (long_short_list[y] == columns[which_data_set][long]):
                    high_vix_long_losses+=1
                elif (long_short_list[y] == columns[which_data_set][short]):
                    high_vix_short_losses+=1
                else:
                    pass
                y+=1
            else:
                if (long_short_list[y] == columns[which_data_set][long]):
                    low_vix_long_losses+=1
                elif (long_short_list[y] == columns[which_data_set][short]):
                    low_vix_short_losses+=1
                else:
                    pass
                y+=1       

for cell in vix_cells:
    try:
        vix_list.append(int(cell.value))
    except:
        vix_list.append(cell.value)
for cell in long_short_cells:
    long_short_list.append(cell.value)

for i in range(12, 21):
    high_vix_long_wins, high_vix_long_losses = 0, 0
    low_vix_long_wins, low_vix_long_losses = 0, 0
    high_vix_short_wins, high_vix_short_losses = 0, 0
    low_vix_short_wins, low_vix_short_losses = 0, 0
    iterables = (win_cells, win_cells_two, loss_cells, loss_cells_two)
    for iterable in iterables:
        countWins(iterable, i)
    print("vix high/low pivot level: %d" % (i))
    try:
        high_vix_long_total = high_vix_long_wins + high_vix_long_losses
        high_vix_long_win_rate = round(high_vix_long_wins/high_vix_long_total * 100, 2)
        print("high vix long win rate: %.2f out of %d trades" % (high_vix_long_win_rate, high_vix_long_total))
    except:
        print("no long above %d" % (i))
    try:
        low_vix_total = low_vix_long_wins + low_vix_long_losses
        low_vix_long_win_rate = round(low_vix_long_wins/low_vix_total* 100, 2)
        print("low vix long win rate: %.2f out of %d trades" % (low_vix_long_win_rate, low_vix_total))
    except:
        print("no longs below %d" % (i))
    try:
        high_vix_short_total = high_vix_short_wins + high_vix_short_losses
        high_vix_short_win_rate = round(high_vix_short_wins/high_vix_short_total * 100, 2)
        print("high vix short win rate: %.2f out of %d trades" % (high_vix_short_win_rate, high_vix_short_total))
    except:
        print("no shorts above %d" % (i))
    try:
        low_vix_short_total = low_vix_short_wins + low_vix_short_losses
        low_vix_short_win_rate = round(low_vix_short_wins/low_vix_short_total * 100, 2)
        print("low vix short win rate: %.2f out of %d trades" % (low_vix_short_win_rate, low_vix_short_total))
    except:
        print("no shorts below %d" % (i))
    print("-------------------------------------------")










