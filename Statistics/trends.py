##find win rates for different ranges of opening price fluct percentages
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
workbook1 = wks.get_worksheet(12)


range_start = "143"
range_end = "648"

win_cells = workbook1.range("O" + range_start + ":O" + range_end)
win_cells_list = []
win_cells_two = workbook1.range("P" + range_start + ":P" + range_end)
win_cells_two_list = []
loss_cells = workbook1.range("Q" + range_start + ":Q" + range_end)
loss_cells_list = []
loss_cells_two = workbook1.range("R" + range_start + ":R" + range_end)
loss_cells_two_list = []
industry_cells = workbook1.range("Y" + range_start + ":Y" + range_end)
industry_list = []
range_trend_cells = workbook1.range("V" + range_start + ":V" + range_end)
range_trend_list = []
long_short_cells = workbook1.range("L" + range_start + ":L" + range_end)
long_short_list = []
opening_flucts = workbook1.range("C" + range_start + ":C" + range_end)
opening_flucts_list = []
trend_long_losses = 0
trend_short_losses = 0
range_long_losses = 0
range_short_losses = 0
trend_long_wins = 0
trend_short_wins = 0
range_long_wins = 0
range_short_wins = 0
opening_fluct_and_trend_wins = 0
opening_fluct_and_trend_losses = 0



def getColNum(cellInfo):
    col = ''
    string = str(cellInfo)
    m = re.search('R(.+?) ', string)
    if (m):
        col = m.group(1)
        return (col[-2:])

#TODO - get count of wins that were long and short


def countWins(iterable):
    global trend_long_losses
    global trend_short_losses
    global range_long_losses
    global range_short_losses
    global trend_long_wins
    global trend_short_wins 
    global range_long_wins
    global range_short_wins
    global opening_fluct_and_trend_wins
    global opening_fluct_and_trend_losses
    min_gap, max_gap = 1, 5
    y = 0
    for cell in iterable:
        if (cell.value == '' or range_trend_list[y] == '' or long_short_list[y] == ''):
            y+=1
            continue
        elif (cell.value != '' and range_trend_list != ''):
            #win columns
            if (getColNum(cell) == "15" or getColNum(cell) == "16"):
                if (range_trend_list[y] == "Range"):
                    if (long_short_list[y] == 'L'):
                        range_long_wins+=1
                    elif (long_short_list[y] == 'S'):
                        range_short_wins+=1
                    else:
                        pass
                    y+=1
                elif (range_trend_list[y] == "Trend"):
                    if ((abs(float(opening_flucts_list[y])) >= min_gap and abs(float(opening_flucts_list[y])) <= max_gap)):
                        opening_fluct_and_trend_wins +=1
                    if (long_short_list[y] == 'L'):
                        trend_long_wins+=1
                    elif (long_short_list[y] == 'S'):
                        trend_short_wins+=1
                    else:
                        pass
                    y+=1
            #loss columns
            else:
                if (range_trend_list[y] == "Range"):
                    if (long_short_list[y] == 'L'):
                        range_long_losses+=1
                    elif (long_short_list[y] == 'S'):
                        range_short_losses+=1
                    else:
                        pass
                    y+=1
                elif (range_trend_list[y] == "Trend"):
                    if ((abs(float(opening_flucts_list[y])) >= min_gap and abs(float(opening_flucts_list[y])) <= max_gap)):
                        opening_fluct_and_trend_losses +=1
                    if (long_short_list[y] == 'L'):
                        trend_long_losses+=1
                    elif (long_short_list[y] == 'S'):
                        trend_short_losses+=1
                    else:
                        pass
                    y+=1
        else:
            y+=1
        
for cell in range_trend_cells:
    range_trend_list.append(cell.value)
for cell in long_short_cells:
    long_short_list.append(cell.value)
for cell in opening_flucts:
    opening_flucts_list.append(cell.value)
iterables = (win_cells, win_cells_two, loss_cells, loss_cells_two)
for iterable in iterables:
    countWins(iterable)

total_trends_longs = trend_long_wins + trend_long_losses
total_trends_shorts = trend_short_wins + trend_short_losses
total_range_longs = range_long_wins + range_long_losses
total_range_shorts = range_short_wins + range_short_losses
total_trends = trend_long_wins + trend_long_losses + trend_short_wins + trend_short_losses
total_ranges = range_long_wins + range_long_losses + range_short_wins + range_short_losses
trend_long_win_rate = round((trend_long_wins/total_trends_longs) * 100, 2)
trend_short_win_rate = round((trend_short_wins/total_trends_shorts) * 100, 2)
range_long_win_rate = round((range_long_wins/total_range_longs) * 100, 2)
range_short_win_rate = round((range_short_wins/total_range_shorts) * 100, 2)
trending_general_win_rate = round(((trend_long_wins + trend_short_wins)/total_trends) * 100, 2)
ranging_general_win_rate = round(((range_long_wins + range_short_wins)/total_ranges) * 100, 2)

total_flucts_trends = opening_fluct_and_trend_losses + opening_fluct_and_trend_wins
flucts_trends_win_rate = round(((opening_fluct_and_trend_wins/total_flucts_trends) * 100), 2)

print("Total number of trades: " + str(total_trends + total_ranges))
print("Total trending trades: " + str(total_trends))
print("Total ranging trades: " + str(total_ranges))
print("-----------------------------------")
print("Ranging trade long wins: " + str(range_long_wins))
print("Trending trade long wins: " + str(trend_long_wins))
print("Ranging trade long losses: " + str(range_long_losses))
print("Trending trade  long losses: " + str(trend_long_losses))
print("Ranging trade short wins: " + str(range_short_wins))
print("Trending trade short wins: " + str(trend_short_wins))
print("Ranging trade short losses: " + str(range_short_losses))
print("Trending trade short losses: " + str(trend_short_losses))
print("-----------------------------------")
print("Trending general win rate: " + str(trending_general_win_rate) + "% ")
print("Ranging general win rate: " + str(ranging_general_win_rate) + "% ")
print("-----------------------------------")
print("Trending long win rate: " + str(trend_long_win_rate) + "% ")
print("Trending short win rate: " + str(trend_short_win_rate) + "% ")
print("Ranging long win rate: " + str(range_long_win_rate) + "% ")
print("Ranging short win rate: " + str(range_short_win_rate) + "% ")
print("total number of trending trades with target opening fluct: %d win rate: %.2f" % (total_flucts_trends, flucts_trends_win_rate))








