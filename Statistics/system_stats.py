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
from bs4 import BeautifulSoup
from oauth2client.service_account import ServiceAccountCredentials
import sys
import time
import subprocess
import re
from functools import reduce
from datetime import datetime
from threading import Thread
import pandas as pd
import requests
directory = "/Users/jakezimmerman/Documents/IBJts2/source/pythonclient2/"
sys.path.insert(0, directory)

scope = ["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/spreadsheets"]
credentials_file = "/Users/jakezimmerman/Documents/python3credentials.json"
credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
gc = gspread.authorize(credentials)
wks = gc.open("RCP90")
data_log = wks.get_worksheet(9)
#system = wks.get_worksheet(0)
range_start = 401
range_end = range_start
ticker_cells = data_log.range('A' + str(range_start) + ":A4000")
for cell in ticker_cells:
    if (cell.value == ''):
        range_end-=1
        break
    range_end+=1

print("range start: {} range end: {}".format(range_start, range_end))

range_start, range_end = str(range_start), str(range_end)
vix_cells = data_log.range("X" + range_start + ":X" + range_end)
vix_list = []
for cell in vix_cells:
    if (cell.value == ''):
        vix_list.append(-1)
    else:
        vix_list.append(int(cell.value))
long_short_cells = data_log.range("I" + range_start + ":I" + range_end)
long_short_list = []
for cell in long_short_cells:
    long_short_list.append(cell.value)
trade_type_cells = data_log.range("T" + range_start + ":T" + range_end)
trade_type_list = []
for cell in trade_type_cells:
    trade_type_list.append(cell.value)
portfolio_win_loss_cells = data_log.range("AG" + range_start + ":AG" + range_end)
portfolio_win_loss_list = []
for cell in portfolio_win_loss_cells:
    portfolio_win_loss_list.append(float(cell.value))
commission_cells = data_log.range("AJ" + range_start + ":AJ" + range_end)
commission_list = []
for cell in commission_cells:
    commission_list.append(float(cell.value))


loss_count, loss_sum = 0, 0
win_count, win_sum = 0, 0
i = 0
for result in portfolio_win_loss_list:
    if (vix_list[i] == -1 or trade_type_list[i] == "crypto"):
        i+=1
        continue
    if (vix_list[i] < 16 and long_short_list[i] == "SELL"):
        i+=1
        continue
    if (result >= 0):
        win_count+=1
        win_sum+=result
    else:
        loss_count+=1
        loss_sum+=result
    i+=1

average_gain = round(win_sum/win_count, 2)
average_loss = round(loss_sum/loss_count, 2)
return_ = win_sum - abs(loss_sum)
print("percentage return {}".format(return_))
average_risk_reward = average_gain/abs(average_loss)
print("average risk reward: ", average_risk_reward)
total_trades = win_count + loss_count
win_rate = round((win_count/total_trades) * 100, 2)
loss_rate = 100 - win_rate
print("win rate: {} out of {} trades".format(win_rate, total_trades))
#expectancy = round(((win_rate/100 * average_risk_reward) - ((loss_rate/100))), 2)
#print("expectancy: ", expectancy)













