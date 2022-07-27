from ibpythonic import ibConnection, message
from ibapi.contract import Contract
from ibapi.order import Order
from ibapi.ticktype import TickTypeEnum as tt
from ibapi.execution import ExecutionFilter
from rate_of_change import rateOfChange
from technical_analysis import TechnicalAnalysis
from time import sleep
import time
import json  #3 imports below are for google sheets api
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from bs4 import BeautifulSoup
import requests
import re
import subprocess
from datetime import datetime
import sys
import ast

account_size = 0
momo_parameters = {"micro": [.35, .45], "small": [.25, .35], "mid": [.20, .30], "large": [.15, .25], "mega": [.15, .25]}
momo_percentage_one, momo_percentage_two = 1, 5
risk_reward_parameters = [3, 1.75, .25, 2, 1.25, .25]
cup_reward, cup_trail_trigger, cup_trail_execute, momo_reward, momo_trail_trigger, momo_trail_execute = 0, 1, 2, 3, 4, 5
over_exhausted = [1.5, 3]
#0 = unit 1, 1 = unit 2, 2 = risk,3 = opening percent
fade_parameters = [.62, .62, 5.5, 20] 
unit_one, unit_two, opening_qualifier_one, opening_qualifier_two = 0, 1, 2, 3
tickerInPlay = ""
executionDict = {}
SYMBOL, SIDE, PRICE, SHARES = 0, 1, 2, 3
screenerDict = {}
OPEN, CLOSE, PERCENTCHANGE, DIRECTION, EXCHANGE, TREND, CLOUDSCORES, EARNINGS, VOLUME = 0, 1, 2, 3, 4, 5, 6, 7, 8
valueDict = {}
LAST, MARKETCAP, LOGGINGID, PRICE2BOOK, PRICEEARNINGS, BETA, STOCH, OUTSIDEHOURSHIGHLOW, PASSBEFORE10, HIGH, LOW = 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10
momoDict = {}
RISK, REWARD, TRAIL_TRIGGER, TRAIL_EXECUTE, PEAK_TROUGH, CHECKING, LASTPEAKTROUGH = 0, 1, 2, 3, 4, 5, 6
optionsDict = {}
EXPIRATIONS, STRIKELIST, CON_ID, OPTIONS_DATA = 0, 1, 2, 3
bid, ask, delta, gamma, theta, vega, vol_, call_open_interest, put_open_interest, current_strike = 0, 1, 2, 3, 4, 5, 6, 7, 8, 9
order_id = 0
temp_id = 0
dataId = 1
populated = False
updated = False
screening = True
override_vix_rule = []
list_of_bars = []
error_list, id_list = [], []
nyse_list = ["TCS", "GPRO", "FIVE", "KEYS", "T", "GDX", "OIH", "SPCE"]
nasdaq_list = ["INTC", "WING", "BLBD", "MATW", "STIM", "SMTC", "CSCO", "MSFT"]
longs_off, shorts_off = False, False
vix_checks = {"09:45" : False, "09:55" : False, "10:10" : False, "10:10" : False, "10:30" : False, "10:50" : False, "11:10" : False}
available_funds = [0, 0]
available, sma = 0, 1
opening, closing, high, low, bar_date, vol, wap, bar_count = 0, 1, 2, 3, 4, 5, 6, 7
price_value, date = 0, 1
offset, log_offset = 3, 2
reset_count = 0
socket_port = 4000
market_data_type = 1
close_status_types = ["win", "loss", "juice"]
#google sheets api linking
credentials_file = "/Users/jakezimmerman/Documents/python3credentials.json"
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
gc = gspread.authorize(credentials)
wks = gc.open("RCP90")
which_sheet, which_screen, which_log = 7, 8, 9
system = wks.get_worksheet(which_sheet)
screener = wks.get_worksheet(which_screen)
data_log = wks.get_worksheet(which_log)
request_options = False #this is here so the marketData callback handler can determine which dictionary to populate

##wrapper functions
#def reply_handler(msg):
    #"""Handles of server replies"""
    #print("Server Response: %s, %s" % (msg.typeName, msg)) 

def errorHandler(msg):
    if ((msg.errorMsg == "No security definition has been found for the request" and msg.id not in error_list) or ((msg.errorCode == 366 or msg.errorCode == 162) and msg.id not in error_list)):
        print("%.1f added to error list" % (msg.id))
        error_list.append(msg.id)
    if (msg.errorCode == 321 or msg.errorCode == 434):
        print("invalid duration for request")
        reConnect()
    if (msg.errorMsg == "Historical Market Data Service error message:Historical data request pacing violation"):
        reConnect()
    print(msg)

def marketData(msg):
    global tickerInPlay, dataId
    ticker = tickerInPlay
    if (msg.reqId == dataId):
        if (ticker in valueDict):
            if (msg.tickType == tt.FUNDAMENTAL_RATIOS):
                find_market_cap = re.search('MKTCAP=(.+?);', msg.value)
                if (find_market_cap):
                    valueDict[ticker][MARKETCAP] = marketCap(float(find_market_cap.group(1)))
                find_p2b = re.search('PRICE2BK=(.+?);', msg.value)
                if (find_p2b):
                    valueDict[ticker][PRICE2BOOK] = find_p2b.group(1)
                find_pe = re.search('APENORM=(.+?);', msg.value)
                if (find_pe):
                    valueDict[ticker][PRICEEARNINGS] = find_pe.group(1)
                find_pe = re.search('BETA=(.+?);', msg.value)
                if (find_pe):
                    valueDict[ticker][BETA] = find_pe.group(1)
        if (not request_options):
            if (ticker in screenerDict):
                if (msg.tickType == tt.OPEN and screenerDict[ticker][OPEN] == 0):
                    #print("ticker: %s open: %.2f" % (ticker, msg.price))
                    screenerDict[ticker][OPEN] = msg.price
                if (msg.tickType == tt.CLOSE and screenerDict[ticker][CLOSE] == 0):
                    #print("ticker: %s close: %.2f" % (ticker, msg.price))
                    screenerDict[ticker][CLOSE] = msg.price
            if (ticker in valueDict):
                if (msg.tickType == tt.LAST):#print("ticker: " + tickerInPlay)
                    valueDict[ticker][LAST] = msg.price
                    #print(" last price " + str(msg.price))
                if (msg.tickType == tt.HIGH):#print("ticker: " + tickerInPlay)
                    valueDict[ticker][HIGH] = msg.price
                    #print(" high price " + str(msg.price))
                if (msg.tickType == tt.LOW):#print("ticker: " + tickerInPlay)
                    valueDict[ticker][LOW] = msg.price
                    #print(" low price " + str(msg.price))
        else:
            if (ticker in optionsDict):
                if (msg.tickType == tt.BID):
                    #print("ask: " +str(msg.price))
                    optionsDict[ticker][OPTIONS_DATA][bid] = msg.price
                    #print("volume: " + str(msg.size))
                if (msg.tickType == tt.ASK):
                    optionsDict[ticker][OPTIONS_DATA][ask] = msg.price
                if (msg.tickType == tt.OPTION_PUT_OPEN_INTEREST):
                    optionsDict[ticker][OPTIONS_DATA][put_open_interest] = msg.size
                if (msg.tickType == tt.OPTION_CALL_OPEN_INTEREST):
                    optionsDict[ticker][OPTIONS_DATA][call_open_interest] = msg.size
                if (msg.tickType == tt.MODEL_OPTION):
                    optionsDict[ticker][OPTIONS_DATA][delta] = abs(round(msg.delta, 3))
                    optionsDict[ticker][OPTIONS_DATA][gamma] = round(msg.gamma, 3)
                    optionsDict[ticker][OPTIONS_DATA][vega] = round(msg.vega, 3)
                    optionsDict[ticker][OPTIONS_DATA][theta] = round(msg.theta, 3)
                if (msg.tickType == tt.VOLUME):
                    optionsDict[ticker][OPTIONS_DATA][vol_] = msg.size

def accountSummary(msg):
    global account_size
    if (msg.tag == "AvailableFunds"):
        available_funds[available] = float(msg.value)
    if (msg.tag == "SMA"):
        available_funds[sma] = float(msg.value)
    if (msg.tag == "NetLiquidation"):
        account_size = float(msg.value)


def securityDefinitionOptionParameter(msg):
    global tickerInPlay
    ticker = tickerInPlay
    optionsDict[ticker][EXPIRATIONS] = msg.expirations
    optionsDict[ticker][STRIKELIST] = msg.strikes

#used to retrieve underlying con id for options
def symbolSamples(msg):
    ticker = tickerInPlay
    for contractDescription in msg.contractDescriptions:
        derivSecTypes = ""
        for derivSecType in contractDescription.derivativeSecTypes:
            derivSecTypes += derivSecType
            derivSecTypes += " "
            if (contractDescription.contract.symbol == ticker):
                #print("conn id: %s" % (contractDescription.contract.conId))
                if ("OPT" in derivSecTypes):
                    optionsDict[ticker][CON_ID] = int(contractDescription.contract.conId)
                #print("Contract: conId:%s, symbol:%s, secType:%s primExchange:%s, currency:%s, derivativeSecTypes:%s" % (contractDescription.contract.conId,
                      #contractDescription.contract.symbol, contractDescription.contract.secType, contractDescription.contract.primaryExchange, contractDescription.contract.currency, derivSecTypes))

def historicalData(msg):
    if (msg.reqId == dataId):
        bar_list = [0, 0, 0, 0, 0, 0, 0, 0]
        #print(type(opening), type(closing), type(high), type(low))
        bar_list[opening], bar_list[closing], bar_list[high], bar_list[low] = msg.bar.open, msg.bar.close, msg.bar.high, msg.bar.low
        bar_list[bar_date], bar_list[vol], bar_list[wap], bar_list[bar_count] = msg.bar.date, msg.bar.volume, msg.bar.average, msg.bar.barCount
        list_of_bars.append(bar_list)

def historicalDataEnd(msg):
    id_list.append(msg.reqId)
    print(msg)

def nextValidId(msg):
    global order_id
    if (order_id == 0):
        order_id = msg.orderId
        print("order id in callback: ", msg.orderId)
    
def execDetails(msg):
    global order_id, temp_id
    if (msg.execution.orderId == order_id):
        if (msg.execution.orderId not in executionDict):
            executionDict.update({msg.execution.orderId : [msg.contract.symbol, msg.execution.side, round(msg.execution.avgPrice, 2), 0]})
        else:
            executionDict[msg.execution.orderId][SHARES]+=msg.execution.shares
    if (msg.execution.orderId == temp_id):
        if (msg.execution.orderId not in executionDict):
            executionDict.update({msg.execution.orderId : [msg.contract.symbol, msg.execution.side, round(msg.execution.avgPrice, 2), 0]})
        else:
            executionDict[msg.execution.orderId][SHARES]+=msg.execution.shares

##script functions
def checkExecutions(open_close, unit, order_id):
    def preMarketPercentage(side, trade_type, bars, entry_price): #this data point needs equity fill price
        def lows():
            prices, low_ = [], 10000000
            for bar in bars:
                prices.append(bar[low])
            for price in prices:
                if (price < low_):
                    low_ = price
            low_to_entry_percentage = round(percentChange(low_, entry_price), 2)
            if (low_ > entry_price):
                low_to_entry_percentage = low_to_entry_percentage * -1
            return(low_to_entry_percentage)
        def highs():
            prices, high_ = [], 0
            for bar in bars:
                prices.append(bar[high])
            for price in prices:
                if (price > high_):
                    high_ = price
            high_to_entry_percentage = round(percentChange(high_, entry_price), 2)
            if (high_ < entry_price):
                high_to_entry_percentage = high_to_entry_percentage * -1 #-1 if entry is in the opposite direction
            return(high_to_entry_percentage)
        if (side == 'L' and trade_type == 'F'): #lows
            return(lows())
        if (side == 'S' and trade_type == 'F'): #highs
            return(highs())
        if (side == 'L' and trade_type == 'T'): #highs
            return(highs())
        if (side == 'S' and trade_type == 'T'): #lows
            return(lows())
    logged = False
    while (logged == False):
        try:
            if (open_close == "open"):
                #get opening fill prices
                i = 0
                symbol = ""
                for cell in fill_list:
                    if (execution_list[i] == ''):
                        break
                    if (cell == 0 and order_id_list[i] in executionDict):
                        symbol = execution_list[i]
                        sheet_index = i + offset
                        #this needs to be where the equity price was at the time of the fill
                        print("order id list index: ", order_id_list[i])
                        print("shares: ", executionDict[order_id_list[i]][SHARES])
                        print("shares in list: ", shares_list[i])
                        if (shares_list[i] == executionDict[order_id_list[i]][SHARES]): #this means we are filled
                            try:
                                requestHistoricalData(symbol, "pre_market", '')
                            except Exception as e:
                                print("pre market historical data error: ", e)
                            if (data_log.acell('AC' + str(valueDict[symbol][LOGGINGID])).value == ''): #selected equity
                                price_ = executionDict[order_id_list[i]][PRICE]
                                system.update_acell('J' + str(sheet_index), price_)
                                fill_list[i] = price_
                            else: #approximate if it selected an option
                                system.update_acell('J' + str(sheet_index), valueDict[symbol][LAST])
                                fill_list[i] = valueDict[symbol][LAST]
                            if (len(list_of_bars) != 0):
                                try:
                                    print("logging distance from pre market high low")
                                    tt = ''
                                    if (screenerDict[symbol][OPEN] > screenerDict[symbol][CLOSE] and screenerDict[symbol][DIRECTION] == 'L'):
                                        tt = 'T'
                                    if (screenerDict[symbol][OPEN] > screenerDict[symbol][CLOSE] and screenerDict[symbol][DIRECTION] == 'S'):
                                        tt = 'F'
                                    if (screenerDict[symbol][OPEN] < screenerDict[symbol][CLOSE] and screenerDict[symbol][DIRECTION] == 'L'):
                                        tt = 'F'
                                    if (screenerDict[symbol][OPEN] < screenerDict[symbol][CLOSE] and screenerDict[symbol][DIRECTION] == 'S'):
                                        tt = 'T'
                                    print("pre market data point params: ", screenerDict[symbol][DIRECTION], tt, "", fill_list[i])
                                    pre_market_high_low_percentage = preMarketPercentage(screenerDict[symbol][DIRECTION], tt, list_of_bars, fill_list[i])
                                    print("pre market high low percentage: ", pre_market_high_low_percentage)
                                    data_log.update_acell('AA' + str(valueDict[symbol][LOGGINGID]), pre_market_high_low_percentage)
                                except Exception as e:
                                    print("pre market data function error: ", e)
                            else:
                                print("no prints pre market for %s" % (symbol))
                            list_of_bars.clear()
                            try:
                                requestHistoricalData(symbol, "atr day", '')
                            except Exception as e:
                                print("atr day data error")
                            if (len(list_of_bars) != 0):
                                try:
                                    print("logging ATR percentage")
                                    opening_gap = screenerDict[symbol][PERCENTCHANGE]
                                    close = screenerDict[symbol][CLOSE]
                                    gap_dollar_amount = close * (opening_gap/100)
                                    prices = []   
                                    for bar in list_of_bars:
                                        prices.append([bar[closing], bar[bar_date], bar[vol], bar[high], bar[low], bar[opening]])
                                    ta = TechnicalAnalysis(prices)
                                    atr = ta.average_true_range(14)[-1][0]
                                    gap_dollar_percentage_atr = round((gap_dollar_amount/atr) * 100, 2)
                                    data_log.update_acell('AL' + str(valueDict[symbol][LOGGINGID]), gap_dollar_percentage_atr)
                                except Exception as e:
                                    print("error logging atr %", e)
                            list_of_bars.clear()
                            try:
                                requestHistoricalData(symbol, "atr min", '')
                            except Exception as e:
                                print("atr day data error")
                            if (len(list_of_bars) != 0):
                                try:
                                    print("logging stop percentage vs atr")
                                    stop_percentage = peak_trough_list[i]/100
                                    fill_price = fill_list[i]
                                    print("fill price: ", fill_price)
                                    stop_dollar_amount = fill_price * stop_percentage
                                    print("stop dollar amount: ", stop_dollar_amount)
                                    prices = []   
                                    for bar in list_of_bars:
                                        prices.append([bar[closing], bar[bar_date], bar[vol], bar[high], bar[low], bar[opening]])
                                    ta = TechnicalAnalysis(prices)
                                    print("displaying average true range")
                                    print(ta.average_true_range(14))
                                    atr = ta.average_true_range(14)[-1][0]
                                    print("5 min atr: ", atr)
                                    stop_dollar_percentage_atr = round((stop_dollar_amount/atr) * 100, 2)
                                    data_log.update_acell('AN' + str(valueDict[symbol][LOGGINGID]), stop_dollar_percentage_atr)
                                except Exception as e:
                                    print("error logging atr min") 
                            list_of_bars.clear()     
                        executionDict[order_id_list[i]][SHARES] = 0
                        #log fills - only the data log will display options fill prices
                        data_log.update_acell('C' + str(valueDict[symbol][LOGGINGID]), round(executionDict[order_id_list[i]][PRICE], 2))
                        logged = True
                    i+=1
                logged = True
            else:
                if (order_id in executionDict):
                    index = int(valueDict[executionDict[order_id][SYMBOL]][LOGGINGID]) - log_offset #offset for data log
                    if (executionDict[order_id][PRICE] != 0):
                        #only for fades as of now because we are experimenting with averaging out of units  
                        if (close_one_list[index] != 0):
                            average_price = (close_one_list[index] + executionDict[order_id][PRICE]) / 2
                            data_log.update_acell('E' + str(index + log_offset), average_price)
                            logged = True
                            close_one_list[index] = average_price
                        else:
                            data_log.update_acell('E' + str(index + log_offset), executionDict[order_id][PRICE])
                            logged = True
                            close_one_list[index] = executionDict[order_id][PRICE]
                logged = True
        except Exception as e:
            print("check executions: " + str(e))
            reConnect()
            continue
            
def close(ticker, ticker_type, side, fill_price, shares, trail_status, index):
    risk, reward, trail_trigger, trail_execute = 0, 1, 2, 3
    if (ticker_type == "fade" and peak_trough_list[index] == 0):
        print("closing fade normal")
        fade(ticker, side, index, "close", fill_price, trail_status)
        return
    else:
        if (valueDict[ticker][MARKETCAP] != '' or ticker_type == "launch"):
            if (ticker_type == "fade"):
                print("closing fade cup")
            elif (ticker_type == "trend"):
                print("closing cup")
            else:
                print("closing momo")
            risk = momoDict[ticker][RISK]
            reward = momoDict[ticker][REWARD]
            trail_trigger = momoDict[ticker][TRAIL_TRIGGER]
            trail_execute = momoDict[ticker][TRAIL_EXECUTE]
        else:
            return
    try:
        price_percentage_change = percentChange(fill_price, valueDict[ticker][LAST])
        if (side == "BOT"):
            #check trail trigger
            if (valueDict[ticker][LAST] > fill_price and price_percentage_change >= trail_trigger and trail_status == ''):
                system.update_acell('K' + str(index + offset), "trail")
                status_list[index] = "trail"
            #win
            if (valueDict[ticker][LAST] > fill_price and price_percentage_change >= reward):
                sendCloseOrder(ticker, "juice", shares, "SELL", index, ticker_type) 
                subprocess.call(["afplay", "barin.mp4"]) #bar and in sound for juice :D
            #trail
            elif (valueDict[ticker][LAST] > fill_price and price_percentage_change <= trail_execute and trail_status == "trail"):
                sendCloseOrder(ticker, "win", shares, "SELL", index, ticker_type) 
            #loss
            elif (valueDict[ticker][LAST] < fill_price and price_percentage_change >= risk):
                sendCloseOrder(ticker, "loss", shares, "SELL", index, ticker_type)
            else:
                pass
                #print("not closed")
        else:
             #check trail trigger
            if (valueDict[ticker][LAST] < fill_price and price_percentage_change >= trail_trigger and trail_status == ''):
                system.update_acell('K' + str(index + offset), "trail")
                status_list[index] = "trail"
            #win
            if (valueDict[ticker][LAST] < fill_price and price_percentage_change >= reward):
                sendCloseOrder(ticker, "juice", shares, "BUY", index, ticker_type) 
                subprocess.call(["afplay", "barin.mp4"])
            #trail
            elif (valueDict[ticker][LAST] < fill_price and price_percentage_change <= trail_execute and trail_status == "trail"):
                sendCloseOrder(ticker, "win", shares, "BUY", index, ticker_type) 
            #loss
            elif (valueDict[ticker][LAST] > fill_price and price_percentage_change >= risk):
                sendCloseOrder(ticker, "loss", shares, "BUY", index, ticker_type)
            else:
                pass
                #print("not closed")
    except Exception as e:
        print("error 2: " + str(e))

#only being used with regular fades for now. Make sure the previous minute bar closed in the right direction
def checkPreviousCandles(ticker, side):
    print("checking previous minute candle")
    try:
        requestHistoricalData(ticker, "candle", '')
    except Exception as e:
        print("candle historical data request error")
    if (len(list_of_bars) != 0):
        prices = []
        for price in list_of_bars:
            prices.append([price[closing], price[bar_date], price[vol], price[high], price[low], price[opening]])
        list_of_bars.clear()
        ta = TechnicalAnalysis(prices) 
        color = 0
        candles = ta.candlesticks()
        previous_candle = candles[1][color]
        current_candle = candles[0][color]
        if (side == "BOT" and previous_candle == "green"):
            return True
        elif (side == "SLD" and previous_candle == "red"):
            return True
        else:
            print("waiting on candle")
            return False
    else:
        print("no data for candle")
        return False


#scalp trade. This is the only trade currently that has a mean reversion risk reward, but a quarter of the position can capture tail risk and is held until day end
def fade(ticker, side, index, phase, fill_price, status): #fill_price is only used when closing. Python allows for optional args
    if (phase == "open"): 
        try: 
            stretch_one = screenerDict[ticker][PERCENTCHANGE] * fade_parameters[unit_one]
            #stretch_two = screenerDict[ticker][PERCENTCHANGE] * fade_parameters[unit_two]
            if (screenerDict[ticker][TREND] == "fade" and screenerDict[ticker][PERCENTCHANGE] < fade_parameters[opening_qualifier_two]):
                print("checking fade entry for: ", ticker, "on side: ", side)
                shares = equityShares(valueDict[ticker][LAST], "fade", ticker, side)
                current_percentage = round(percentChange(screenerDict[ticker][OPEN], valueDict[ticker][LAST]), 2)
                if (side == "BOT" and (not longs_off or ticker in override_vix_rule)): #longing fades requires a bigger earnings surprise when there is no uptrend, but the check is done earlier on.
                    print("current fade percentage: ", current_percentage)
                    if (valueDict[ticker][LAST] < screenerDict[ticker][OPEN] and percentChange(screenerDict[ticker][OPEN], valueDict[ticker][LAST]) >= stretch_one and ticker not in execution_list):
                        if (checkPreviousCandles(ticker, side)):
                            print("sending fade long order")
                            sendNewOrder(ticker, shares, "BUY", "one", index, "fade")
                if (side == "SLD" and (not shorts_off or ticker in override_vix_rule)):
                    print("current fade percentage: ", current_percentage)
                    if (valueDict[ticker][LAST] > screenerDict[ticker][OPEN] and percentChange(screenerDict[ticker][OPEN], valueDict[ticker][LAST]) >= stretch_one and ticker not in execution_list):
                        if (checkPreviousCandles(ticker, side)):
                            print("sending fade short order")
                            sendNewOrder(ticker, shares, "SELL", "one", index, "fade")
            else:
                print("passing fade")
                pass
        except Exception as e:
            print("fade open: " + str(e))
    else:
        print("fade close for %s, side: %s, fill: %.2f, status: %s" % (ticker, side, fill_price, status))
        try:    
            #if stoploss is changed, it needs to also be changed in equityShares
            shares = shares_list[index]/2
            #trail = screenerDict[ticker][PERCENTCHANGE]/24
            #trail = .25
            exit_half = screenerDict[ticker][PERCENTCHANGE]/1.3
            exit_full = screenerDict[ticker][PERCENTCHANGE]/1.2
            stop_loss = screenerDict[ticker][PERCENTCHANGE]/1.5
            current_percentage = percentChange(fill_price, valueDict[ticker][LAST])
            if (side == "BOT"):
                stop_price = fill_price - (fill_price * (stop_loss/100))
                print("stoploss: ", stop_price)
                if (valueDict[ticker][LAST] > fill_price and current_percentage >= exit_half and status == ''):
                    sendCloseOrder(ticker, "half", shares, "SELL", index, '')
                elif (valueDict[ticker][LAST] <= fill_price and status == "half"):
                    sendCloseOrder(ticker, "win", shares, "SELL", index, '')
                elif (valueDict[ticker][LAST] > fill_price and current_percentage >= exit_full):
                    sendCloseOrder(ticker, "juice", shares, "SELL", index, '')
                    subprocess.call(["afplay", "barin.mp4"])
                elif (valueDict[ticker][LAST] < fill_price and current_percentage >= stop_loss):
                    sendCloseOrder(ticker, "loss", shares * 2, "SELL", index, '') #multiple shares by 2 because stop loss gets out of full unit
                else:
                    pass
            else:
                stop_price = fill_price + (fill_price * (stop_loss/100))
                print("stoploss: ", stop_price)
                if (valueDict[ticker][LAST] < fill_price and current_percentage >= exit_half and status == ''):
                    sendCloseOrder(ticker, "half", shares, "BUY", index, '')
                elif (valueDict[ticker][LAST] >= fill_price and status == "half"):
                    sendCloseOrder(ticker, "win", shares, "BUY", index, '')
                elif (valueDict[ticker][LAST] < fill_price and current_percentage >= exit_full):
                    sendCloseOrder(ticker, "juice", shares, "BUY", index, '')
                    subprocess.call(["afplay", "barin.mp4"])
                elif (valueDict[ticker][LAST] > fill_price and current_percentage >= stop_loss):
                    sendCloseOrder(ticker, "loss", shares * 2, "BUY", index, '')
                else:
                    pass
        except Exception as e:
            print("fade close: " + str(e))

def execution(ticker, recent_high_low, ticker_type, side, index):
    k_value, d_value = 0, 1
    concavity, peak_trough = 0, 1
    unit_one, unit_two, pullback = 0, 1, 0
    shares = 0

    def setParameters(risk_, reward_, peak_trough_, trail_trigger_, trail_execute_):
        momoDict[ticker][RISK] = risk_
        momoDict[ticker][REWARD] = reward_
        momoDict[ticker][PEAK_TROUGH] = peak_trough_
        momoDict[ticker][TRAIL_TRIGGER] = trail_trigger_
        momoDict[ticker][TRAIL_EXECUTE] = trail_execute_

    if (ticker_type == "launch" and ticker not in execution_list):
        side_ = ''
        print("recent high low {} last price {} index {}".format(recent_high_low, valueDict[ticker][LAST], index))
        risk_ = round(percentChange(float(recent_high_low), valueDict[ticker][LAST]), 2) #recent_high_low is the stop price for launchpad
        if (side == "BOT"):
            side_ = "BUY"
            setParameters(risk_, risk_ * risk_reward_parameters[momo_reward], valueDict[ticker][LAST] - (valueDict[ticker][LAST] * (risk_/100)), risk_ * risk_reward_parameters[momo_trail_trigger], risk_ * risk_reward_parameters[momo_trail_execute])
        else:
            side_ = "SELL"
            setParameters(risk_, risk_ * risk_reward_parameters[momo_reward], valueDict[ticker][LAST] + (valueDict[ticker][LAST] * (risk_/100)), risk_ * risk_reward_parameters[momo_trail_trigger], risk_ * risk_reward_parameters[momo_trail_execute])
        shares = equityShares(valueDict[ticker][LAST], ticker_type, ticker, side)
        sendNewOrder(ticker, shares, side_, "one", index, ticker_type)
        return
    #fade/trend executions can only happen day of screening - open/close will be empty for trend/fades submitted in days prior
    if (screenerDict[ticker][TREND] == ''):
        return
    if (ticker_type == "fade" and fades_on): 
        print("ticker value: %s" % (ticker))
        fade(ticker, side, index, "open", 0, "") #return statement has been removed to allow for cup/fade checks for fades
    if (ticker_type == "fade" and not fades_on):
        return
    if (not momos_on): #momos account for fade cups, which will also turn off at 11:30am
        return
    else:
        #pull percentage pullback from appropriate dictionary. pullback percentages are somewhat arbitrary and based on previous moves
        if (valueDict[ticker][MARKETCAP] != ""):
            pullback = momo_parameters[valueDict[ticker][MARKETCAP]][unit_one]
        else:
            print("no market cap available")
            return
    #functions below are only used in execution
    def momoStochRsi():
        try:
            print("requesting stoch rsi for momo entry")
            requestHistoricalData(ticker, "stoch_rsi", '')
            stoch_rsi_prices = []
            for bar in list_of_bars:
                stoch_rsi_prices.append([bar[closing], bar[bar_date]])
            list_of_bars.clear()
            technical_analysis = TechnicalAnalysis(stoch_rsi_prices)
            k_values = technical_analysis.stochastic_rsi(14, 3, 3)[k_value] #arbitrarily using K instead of D
            d_values = technical_analysis.stochastic_rsi(14, 3, 3)[d_value] #arbitrarily using K instead of D
            k = round(k_values[-1][price_value], 2)
            d = round(d_values[-1][price_value], 2)
            print("stoch values k and d: ", k, d)
            print("displaying momo stoch rsi k values")
            print(k_values)
            #most recent 3 rsi bars are below 30 for long, or above 70 for short. This signifies a pause in momentum
            in_target_zone = True
            length_of_k_values = len(k_values)
            if (length_of_k_values == 0):
                print("stoch rsi list empty")
                in_target_zone = False
            if (side == "BOT"):
                for price in k_values[length_of_k_values-3:]:
                    if (price[price_value] > 30):
                        in_target_zone = False
                        break
                if (k <= d): #k must be greater than d to go long and vice versa for short
                    in_target_zone = False
            else:
                for price in k_values[length_of_k_values-3:]:
                    if (price[price_value] < 70):
                        in_target_zone = False
                        break
                if (k >= d):
                    in_target_zone = False
            print("is in target zone: ", in_target_zone)
            return(in_target_zone)
        except Exception as e:
            print("momoStochRsi error: ", e)

    def cupRateOfChange():
        print("requesting historicals for roc")
        requestHistoricalData(ticker, "roc", index)
        prices, highs, lows = [], [], []
        for bar in list_of_bars:
            prices.append(bar[closing])
            highs.append(bar[high])
            lows.append(bar[low])
        list_of_bars.clear()
        print("cup roc prices")
        print(prices)
        print("highs")
        print(highs)
        print("lows")
        print(lows)
        cup_shape = rateOfChange(prices, highs, lows, recent_high_low, side, valueDict[ticker][LAST])
        print("cup status for %s" % (ticker))
        cup_result = cup_shape.getResult()
        c, p_t = cup_result[concavity], cup_result[peak_trough]
        print(c, p_t)
        return([c, p_t])

    #for data log
    def outsideHoursHighLow(): 
        try:
            print("retrieving outside hours high low")
            requestHistoricalData(ticker, "pre_market", '')
            if (side == "BOT"):
                high_ = 0
                for bar in list_of_bars:
                    if (bar[high] > high_):
                        high_ = bar[high]
                list_of_bars.clear()
                return(high_)
            else:
                low_ = 10000000
                for bar in list_of_bars:
                    if (bar[low] < low_):
                        low_ = bar[low]
                return(low_)
        except Exception as e:
            print("outside hours high low error: ", e)

    min_peak_trough = momo_parameters[valueDict[ticker][MARKETCAP]][1]
    cup_concavity = False
    print("ticker: %s pullback:  %.2f high_low %.2f side: %s" % (ticker, pullback, recent_high_low, side))
    if (recent_high_low == -1):
        print("only normal fade allowed")
        return
    #check order conditions
    if (side == "BOT" and (not longs_off or ticker in override_vix_rule)):
        print("checking for long order")
        if (valueDict[ticker][LAST] < recent_high_low and percentChange(valueDict[ticker][LAST], recent_high_low) >= pullback and ticker not in execution_list):
            print("checking long execution for %s of type %s" % (ticker, ticker_type))
            if (not ticker_type == "momo" and cups_on):
                rate_of_change_results = cupRateOfChange()
                print("concavity: ", rate_of_change_results[concavity])
                cup_trough, cup_concavity = rate_of_change_results[peak_trough], rate_of_change_results[concavity]
                if (cup_trough != momoDict[ticker][LASTPEAKTROUGH]): #has the peak changed? in this case, have we made a new low?
                    momoDict[ticker][CHECKING] = True
                    momoDict[ticker][LASTPEAKTROUGH] = cup_trough
            if (not cup_concavity):
                momoDict[ticker][CHECKING] = False
            if (not momoDict[ticker][CHECKING]):
                if (ticker_type == "fade"):
                    return
                try:
                    print("checking long momo criteria")
                    #first if statement should check if pre market high low was passed by 10am
                    #get premarket high low and store in valuedict if its not already there (not databasing in sheet since this high/low happened before market open)
                    outside_hours_high = valueDict[ticker][OUTSIDEHOURSHIGHLOW]
                    if (outside_hours_high == 0):
                        outside_hours_high = outsideHoursHighLow()
                        list_of_bars.clear()
                        valueDict[ticker][OUTSIDEHOURSHIGHLOW] = outside_hours_high
                    print("outside hours high: ", outside_hours_high)
                    #date has to be between open and 10
                    passed_before_10am = valueDict[ticker][PASSBEFORE10]
                    if (passed_before_10am == 0):
                        #get high between 9:30 and 10 or between current time and 9:30 if before 10
                        requestHistoricalData(ticker, "momo", -2)
                        high_ = 0
                        for bar in list_of_bars:
                            if (bar[high] > outside_hours_high):
                                valueDict[ticker][PASSBEFORE10] = True
                                passed_before_10am = True
                        list_of_bars.clear()
                    print("passed before 10: ", passed_before_10am)
                    addendum_percentage = percentChange(screenerDict[ticker][OPEN], valueDict[ticker][HIGH])
                    print("addendum percentage: ", addendum_percentage)
                    if (valueDict[ticker][LAST] > outside_hours_high and passed_before_10am and addendum_percentage > screenerDict[ticker][PERCENTCHANGE] * .75):
                        #get rsi first
                        print("checking long stoch rsi")
                        stoch_rsi = momoStochRsi()
                        if (stoch_rsi):
                            requestHistoricalData(ticker, "momo", -1)
                            prices, high_ = [], 0
                            for bar in list_of_bars:
                                prices.append(bar[closing])
                                if (bar[high] > high_):
                                    high_ = bar[high]
                            list_of_bars.clear()
                            try: #just observing for now
                                momo_roc = rateOfChange(prices, [], [], 0, "SLD", 0) 
                                momo_arc = momo_roc.calcFirstHalf(prices, 0, "SLD")
                                print("momo concavity: ", momo_arc)
                            except Exception as e:
                                print("momo arc error: ", e)
                            print("sending momo long trade")
                            if (ticker_type != "momo"):
                                ticker_type = "momo"
                            risk_ = round(percentChange(high_, valueDict[ticker][LAST]), 2)
                            setParameters(risk_, risk_ * risk_reward_parameters[momo_reward], valueDict[ticker][LAST] - (valueDict[ticker][LAST] * (risk_/100)), risk_ * risk_reward_parameters[momo_trail_trigger], risk_ * risk_reward_parameters[momo_trail_execute])
                            shares = equityShares(valueDict[ticker][LAST], ticker_type, ticker, side)
                            sendNewOrder(ticker, shares, "BUY", "one", index, ticker_type)
                            return
                    else:
                        print("addendum or passed before 10 not satisfied")
                except Exception as e:
                    print("momo long trade error: ", e)
                return
            if (float(screenerDict[ticker][EARNINGS]) >= 0 and screenerDict[ticker][VOLUME] != ''):
                print("checking cup criteria")
                if (cup_concavity):
                    print("long cup trough", cup_trough)
                    momoDict[ticker][RISK] = round(percentChange(valueDict[ticker][LAST], cup_trough), 2)
                    if (momoDict[ticker][RISK] < min_peak_trough):
                        print("trough is too close")
                        momoDict[ticker][RISK] = 0
                        return
                    print("sending cup long trade")
                    setParameters(momoDict[ticker][RISK], momoDict[ticker][RISK] * risk_reward_parameters[cup_reward], cup_trough, momoDict[ticker][RISK] * risk_reward_parameters[cup_trail_trigger], momoDict[ticker][RISK] * risk_reward_parameters[cup_trail_execute])
                    shares = equityShares(valueDict[ticker][LAST], ticker_type, ticker, side)
                    sendNewOrder(ticker, shares, "BUY", "one", index, ticker_type)
                    return
            else:
                if (screenerDict[ticker][VOLUME] == ''):
                    print("no pre market volume for cup")
                else:
                    print("fundamental score is off")
        else:
            print("no order")
    if (side == "SLD" and (not shorts_off or ticker in override_vix_rule)):
        print("checking for short order")
        if (valueDict[ticker][LAST] > recent_high_low and percentChange(valueDict[ticker][LAST], recent_high_low) >= pullback and ticker not in execution_list):
            print("checking short execution for %s of type %s" % (ticker, ticker_type))
            if (not ticker_type == "momo" and cups_on):
                rate_of_change_results = cupRateOfChange()
                cup_peak, cup_concavity = rate_of_change_results[peak_trough], rate_of_change_results[concavity]
                if (cup_peak != momoDict[ticker][LASTPEAKTROUGH]):
                    momoDict[ticker][CHECKING] = True
                    momoDict[ticker][LASTPEAKTROUGH] = cup_peak
            if (not cup_concavity):
                momoDict[ticker][CHECKING] = False
            if (not momoDict[ticker][CHECKING]):
                if (ticker_type == "fade"):
                    return
                try:
                    print("checking short momo criteria")
                    #first if statement should check if pre market high low was passed by 10am
                    #get premarket high low and store in valuedict if its not already there (not databasing in sheet since this high/low happened before market open)
                    outside_hours_low = valueDict[ticker][OUTSIDEHOURSHIGHLOW]
                    if (outside_hours_low == 0):
                        outside_hours_low = outsideHoursHighLow()
                        list_of_bars.clear()
                        valueDict[ticker][OUTSIDEHOURSHIGHLOW] = outside_hours_low
                    print("outside hours low: ", outside_hours_low)
                    #date has to be between open and 10
                    passed_before_10am = valueDict[ticker][PASSBEFORE10]
                    if (passed_before_10am == 0):
                        #get high between 9:30 and 10 or between current time and 9:30 if before 10
                        requestHistoricalData(ticker, "momo", -2)
                        low_ = 1000000
                        for bar in list_of_bars:
                            if (bar[low] < outside_hours_low):
                                valueDict[ticker][PASSBEFORE10] = True
                                passed_before_10am = True
                        list_of_bars.clear()
                    print("passed before 10: ", passed_before_10am)
                    addendum_percentage = percentChange(screenerDict[ticker][OPEN], valueDict[ticker][LOW])
                    print("addendum percentage: ", addendum_percentage)
                    if (valueDict[ticker][LAST] < outside_hours_low and passed_before_10am and addendum_percentage > screenerDict[ticker][PERCENTCHANGE] * .75):
                        #get rsi first
                        print("checking short stoch rsi")
                        stoch_rsi = momoStochRsi()
                        if (stoch_rsi):
                            #check concavity - this measure is the reverse of what is returned by calcFirstHalf in rateOfChangeObject
                            requestHistoricalData(ticker, "momo", -1)
                            prices, low_ = [], 1000000
                            for bar in list_of_bars:
                                prices.append(bar[closing])
                                if (bar[low] < low_):
                                    low_ = bar[low]
                            list_of_bars.clear()
                            try:
                                momo_roc = rateOfChange(prices, [], [], 0, "BOT", 0) 
                                momo_arc = momo_roc.calcFirstHalf(prices, 0, "BOT")
                                print("momo concavity: ", momo_arc)
                            except Exception as e:
                                print("momo arc error: ", e)
                            print("sending momo short trade")
                            if (ticker_type != "momo"):
                                ticker_type = "momo"
                            risk_ = round(percentChange(low_, valueDict[ticker][LAST]), 2)
                            setParameters(risk_, risk_ * risk_reward_parameters[momo_reward], valueDict[ticker][LAST] + (valueDict[ticker][LAST] * (risk_/100)), risk_ * risk_reward_parameters[momo_trail_trigger], risk_ * risk_reward_parameters[momo_trail_execute])
                            shares = equityShares(valueDict[ticker][LAST], ticker_type, ticker, side)
                            sendNewOrder(ticker, shares, "SELL", "one", index, ticker_type)
                    else:
                        print("addendum or passed before 10 not satisfied")
                except Exception as e:
                    print("momo short trade error: ", e)
                return
            if (float(screenerDict[ticker][EARNINGS]) <= 0 and screenerDict[ticker][VOLUME] != ''):
                if (cup_concavity):
                    print("short cup peak", cup_peak)
                    momoDict[ticker][RISK] = round(percentChange(valueDict[ticker][LAST], cup_peak), 2)
                    if (momoDict[ticker][RISK] < min_peak_trough):
                        print("peak is too close")
                        momoDict[ticker][RISK] = 0
                        return
                    print("sending cup short trade")
                    setParameters(momoDict[ticker][RISK], momoDict[ticker][RISK] * risk_reward_parameters[cup_reward], cup_peak, momoDict[ticker][RISK] * risk_reward_parameters[cup_trail_trigger], momoDict[ticker][RISK] * risk_reward_parameters[cup_trail_execute])
                    shares = equityShares(valueDict[ticker][LAST], ticker_type, ticker, side)
                    sendNewOrder(ticker, shares, "SELL", "one", index, ticker_type)
            else:
                if (screenerDict[ticker][VOLUME] == ''):
                    print("no pre market volume for cup")
                else:
                    print("fundamental score is off")
        else:
            print("no order")

#calculate #of minutes that have elapsed between time of check and the open
def calculateRequestTime(index):
    try:
        datetimeFormat = '%Y-%m-%d %H:%M:%S'
        if (index != -1 and index != -2): #cup
            open_ = str(time_volume_list[index][:19])
            now_ = str(datetime.now())[:19]
        elif (index == -1):
            open_ = str(datetime.now())[:11] + "09:30:00"
            now_ = str(datetime.now())[:19]
        else: #momo
            open_ = str(datetime.now())[:11] + "09:30:00"
            now_ = str(datetime.now())[:19]
            if (str(datetime.now())[12] != '9'):
                now_ = str(datetime.now())[:11] + "10:00:00"
        print("open ", open_)
        print("now ", now_)
        date1 = datetime.strptime(open_, datetimeFormat)
        date2 = datetime.strptime(now_, datetimeFormat)
        diff = date1 - date2
        seconds, days = diff.seconds, diff.days
        minutes = (seconds % 3600) // 60
        total_mins = (diff.days*1440 + diff.seconds/60)
        #print("total mins: %.2f" % (total_mins))
        return(abs(int(total_mins)))
    except Exception as e:
        print("calc request time error: ", e)

#return duration in accordance with step size rules depending on time delta (in minutes)
def findDuration(delta):
    seconds = delta * 60
    return(str(seconds) + " S")

def requestHistoricalData(ticker, which_phase, index): 
    global dataId
    try:
        date = str(datetime.now())
        new_date = date[0:10].replace('-', '')
        end_date = new_date + date[10:19]
        exchange = getExchange(ticker)
        contract = createContract(ticker, 'STK', exchange, exchange, 'USD', "", "", "")
        if (which_phase == "roc" or which_phase == "momo"):
            duration = findDuration(calculateRequestTime(index))
            print("roc duration: ", duration)
            print("requesting prices for rate of change")
            if (duration != "0 S"):
                tws.reqHistoricalData(dataId, contract, end_date, duration, "15 secs", "TRADES", 1, 1, False, [])
            else:
                return
        elif (which_phase == "atr day"):
            print("requesting data for day atr")
            tws.reqHistoricalData(dataId, contract, end_date, "1 Y", "1 day", "TRADES", 1, 1, False, [])
        elif (which_phase == "atr min"):
            print("requesting data for min atr")
            tws.reqHistoricalData(dataId, contract, end_date, "2 D", "5 mins", "TRADES", 1, 1, False, [])
        elif (which_phase == "pre_market"):
            end_date = (str(datetime.now())[:10] + " 09:30:00").replace('-', '')
            tws.reqHistoricalData(dataId, contract, end_date, "28800 S", "1 min", "TRADES", 0, 1, False, [])
        elif (which_phase == "stoch_rsi"):
            print("requesting stoch rsi for momo")
            tws.reqHistoricalData(dataId, contract, end_date, "28800 S", "1 min", "TRADES", 1, 1, False, [])
        elif (which_phase == "candle"):
            tws.reqHistoricalData(dataId, contract, end_date, "120 S", "1 min", "TRADES", 1, 1, False, [])
        else:
            return
        wait()
        dataId+=1
    except Exception as e:
        print("request historical data error: ", e)

#equity position sizing 
def equityShares(sharePrice, ticker_type, ticker, side):
    try: 
        stop_loss = equity_max_risk #stop loss dollar amount per trade for stock cup trades
        sharesAmount, risk = 0, 2             
        if (ticker_type == "fade"):
            if (ticker in screenerDict and screenerDict[ticker][PERCENTCHANGE] != 0):
                if (ticker in momoDict and momoDict[ticker][PEAK_TROUGH] != 0):
                    sharesAmount = round(stop_loss/(sharePrice * (momoDict[ticker][RISK]/100)), 0)
                else:#normal fade
                    #stop_loss = stop_loss * 1.5 #shooting for risk in between normal fade option and cup/momo
                    stop_price = 0
                    sharesAmount = round(stop_loss/(sharePrice * ((screenerDict[ticker][PERCENTCHANGE]/1.5)/100)), 0)
                    if (side == "BOT"):
                        stop_price = sharePrice - (sharePrice * ((screenerDict[ticker][PERCENTCHANGE]/1.5)/100))
                    else:
                        stop_price = sharePrice + (sharePrice * ((screenerDict[ticker][PERCENTCHANGE]/1.5)/100))
                    print("stop price for %s: %.2f" % (ticker, stop_price))
        elif (ticker_type == "trend" or ticker_type == "momo" or ticker_type == "launch"):
            print("share price {} momodict {}".format(sharePrice, momoDict[ticker][RISK]))
            sharesAmount = round(stop_loss/(sharePrice * (momoDict[ticker][RISK]/100)), 0) 
        else:
            print("equity shares error")
        if (sharesAmount % 2 == 0):
            return sharesAmount
        else:
            return (sharesAmount - 1)
    except Exception as e:
        print("equity shares error: " + str(e))

#returns position size for cup option trades. size is structured so that ~1-2% of portfolio is risked if cup breaks and peak/trough is reached.
def calculateOptionDollarRisk(delta, gamma, equity_open, equity_stop, option_price): #launchpad motherfucker
    #make sure that delta is an absolute value
    #if equity difference is less than $1, gamma is not needed
    difference = abs(equity_open - equity_stop)
    print("difference: ", difference)
    #get option premium move
    if (difference <= 1):
        premium_lost = 0
        premium_lost = round(delta*difference, 2)
        print("premium lost: ", premium_lost)
        premium_percentage_lost = round(premium_lost/option_price, 2)
        print("premium percentage lost: ", premium_percentage_lost)
        max_dollar_risk = round(cup_max_risk/premium_percentage_lost, 0)
        print("max_dollar_risk: ", max_dollar_risk)
        return(max_dollar_risk)
    else:
         estimated_delta = delta
         anticipated_premium_at_loss = option_price
         #find number of whole dollars in difference and add 1
         dollars = int(difference + 1)
         print("dollars: ", dollars)
         equity_price = equity_open
         i = 1
         for dollar in range(dollars):
            print("iteration: ", i, ' ', anticipated_premium_at_loss)
            if (i < dollars):
                anticipated_premium_at_loss-=estimated_delta
                if ((estimated_delta - gamma) > 0):
                    estimated_delta-=gamma
                equity_price-=1.00
            else:
                print("estimated delta: ", estimated_delta)
                remainder = abs(equity_price - equity_stop)
                print("current premium at loss: ", anticipated_premium_at_loss)
                print("remainder: ", remainder)
                anticipated_premium_at_loss-=(estimated_delta*remainder)
            i+=1
         print("after loop premium value: ", anticipated_premium_at_loss)
         #find percentage loss premium represents against the iitial option price
         premium_percentage_lost = round((option_price - anticipated_premium_at_loss)/option_price, 2)
         print("premium percentage lost: ", premium_percentage_lost)
         max_dollar_risk = round(cup_max_risk/premium_percentage_lost, 0)
         print("max_dollar_risk: ", max_dollar_risk)
         return(max_dollar_risk)

#structure order contains a separate position sizer. 
def structureOrder(ticker, side, sec_type, quantity, contract_info, index):
    global order_id, request_options
    def maxPositionSize(shares, price): #return true if size is acceptable
        position_size = abs(shares * price)
        if (position_size > account_size * 1.5):
            print("position is too big: ", position_size)
            return False
        else:
            return True
    #get number of options contracts
    price, bid_, ask_ = 0, 0, 0
    if (sec_type == "OPT"):
        bid_, ask_ = optionsDict[ticker][OPTIONS_DATA][bid], optionsDict[ticker][OPTIONS_DATA][ask]
        price = bid_
        if (price == -1 or price == 0): #generally, this just means that there is no value in TWS
            print("options price error")
            return(-1)
        quantity = 0
        option_dollar_amount = price * 100
        quantity = int(max_option_cost/option_dollar_amount)
        if (ticker in momoDict and momoDict[ticker][PEAK_TROUGH] != 0):
            price = ask_
            if (price == -1 or price == 0):
                print("options price error")
                return(-1)
            option_dollar_amount = price * 100
            equity_open = valueDict[ticker][LAST]
            equity_stop = momoDict[ticker][PEAK_TROUGH]
            print("equity stop: ", equity_stop)
            quantity = int(calculateOptionDollarRisk(optionsDict[ticker][OPTIONS_DATA][delta], optionsDict[ticker][OPTIONS_DATA][gamma], equity_open, equity_stop, price)/option_dollar_amount)
        if (quantity % 2 != 0):#even number of contracts
            quantity-=1
        if (quantity == 0):
            print("can't place order with 0 contracts!")
            return(-2)
        print("number of contracts: ", quantity)
        if (quantity > 600):
            print("too many contracts. Spread widening risk.")
            return(-2)
        if (quantity < 6):
            print("too few contracts.")
            return(-2)
        d = round(ask_ - bid_, 2) #spread dollar difference
        if (d > .03):
            mid_point = round(bid_ + (d/2), 2)
            price = mid_point
            print("midpoint: ", mid_point)
        #cap max fee at ~ .2% of account. 
        opt_commission = .65 * quantity #IBKR charges .65 per contract.
        fee = (quantity * d * 100) + opt_commission 
        print("quantity: ", quantity, "spread dollar amount: ", d, "commission: ", opt_commission, "fee: ", fee) 
        if (fee < account_size * .001): #take the ask
            price = ask_
        if (fee > account_size * .002): #play stock instead
            print("fee is too high to play option")
            return(-2)
        if (maxPositionSize(quantity, option_dollar_amount)):
            order = createOrder("LMT", quantity, "BUY", price)
            tws.placeOrder(order_id, contract_info, order)
            sleep(.1)
            print("new order placed for %s" % (ticker))
            return(quantity)
        else:
            return(-1)
    else:
        if (side == "BUY"):
            price = round(valueDict[ticker][LAST] + (valueDict[ticker][LAST] * .0075), 1)
            if (maxPositionSize(quantity, price)):
                print("stock buy order")
                order = createOrder("LMT", quantity, side, price)
                tws.placeOrder(order_id, contract_info, order)
                sleep(.1)
                print("new order placed for %s" % (ticker))
            else:
                return(-1)
        else:
            price = round(valueDict[ticker][LAST] - (valueDict[ticker][LAST] * .0075), 1)
            if (maxPositionSize(quantity, price)):
                print("stock short order")
                order = createOrder("LMT", quantity, side, price)
                tws.placeOrder(order_id, contract_info, order)
                sleep(.1)
                print("new order placed for %s" % (ticker))
            else:
                return(-1)

#composes new order to trade options if they are liquid and available, or stock if they are not
def sendNewOrder(ticker, shares, side, unit, index, ticker_type):
    global dataId, request_options
    def send(ticker, shares, side, unit, index, ticker_type, instrument, contract_info): 
        global order_id, while_iteration
        try:
            shares = shares
            print("stock shares: ", shares)
            order_ = structureOrder(ticker, side, instrument, shares, contract_info, index)
            if (order_ == -1):
                if (ticker in momoDict):
                    momoDict.update({ticker : [0, 0, 0, 0, 0, True, 0]})
                return
            if (order_ == -2):
                send(ticker, shares, side, unit, index, ticker_type, "STK", contract) #use global contract for stock
                return
            if (ticker_type == "momo"):
                system.update_acell('D' + str(index + offset), 'M')
                trade_type_list[index] = 'M'
            logData("open", ticker, unit, side, index, "", ticker_type)
            if (instrument == "OPT"):
                shares = order_#structure order returns the quantity (# of contracts) for options
                printExecutions(ticker, shares, order_id, contract_info.lastTradeDateOrContractMonth, contract_info.strike, contract_info.right)
            else:
                printExecutions(ticker, shares, order_id, "", "", "")
            tws.reqExecutions(10001, ExecutionFilter())
            sleep(.1)
            print("unit before check exections ", unit)
            checkExecutions("open", unit, order_id)
            for key, value in executionDict.items():
                print(key, value)
            order_id+=1
            while_iteration+=1 #one execution takes time, so wait until next loop for next set of checks so that the delay for closing out isnt too long
            #print(recent_highs_lows_list[index])
        except Exception as e:
            print("send error: ", e)
    try:
        tws.reqMatchingSymbols(dataId, ticker)
        waiting = True
        start = time.time()
        while (optionsDict[ticker][CON_ID] == 0 and waiting):
            end = time.time()
            if (end - start > timeout):
                waiting = False
        dataId+=1
        print("underlying con id ", optionsDict[ticker][CON_ID])
        tws.reqSecDefOptParams(dataId, ticker, "", "STK", optionsDict[ticker][CON_ID])
        start = time.time()
        waiting = True
        start = time.time()
        while ((optionsDict[ticker][EXPIRATIONS] == 0 or optionsDict[ticker][STRIKELIST] == 0) and waiting):
            end = time.time()
            if (end - start > timeout):
                waiting = False
        dataId+=1
        date_string = str(datetime.now())[:10]
        current_date = date_string.replace('-', '')
        min_days_till_expiration, max_days_till_expiration = 2, 10
        expirations_, chosen_expiration = [], ""
        found_option = False
        try:
            returned_expirations = list(optionsDict[ticker][EXPIRATIONS])
            for exp in returned_expirations:
                expirations_.append(exp.replace('-', ''))
            format_ = "%Y%m%d"
            #print(expirations_)
            expirations_.sort(key=lambda date: datetime.strptime(date, format_))
            print(expirations_)
            for exp in expirations_:
                if (found_option):
                    break
                print("expiration: ", exp)
                days = findNumberOfDays(current_date, exp)
                print("days till expiration: ", days)
                if (days < min_days_till_expiration or days > max_days_till_expiration):
                    continue
                chosen_expiration = exp
                if (chosen_expiration == ""):
                    print("no options available")
                    send(ticker, shares, side, unit, index, ticker_type, "STK", contract)
                else:
                    expiration = chosen_expiration
                    print("expiration: ", expiration)
                    returned_strikes = list(optionsDict[ticker][STRIKELIST])
                    if (side == "BUY"):
                        returned_strikes = sorted(returned_strikes, key=float, reverse=False)
                    else:
                        returned_strikes = sorted(returned_strikes, key=float, reverse=True)
                    print("strikes")
                    print(returned_strikes)
                    for strike in returned_strikes:
                        optionsDict[ticker][OPTIONS_DATA][delta] = 0
                        optionsDict[ticker][OPTIONS_DATA][bid] = 0
                        optionsDict[ticker][OPTIONS_DATA][ask] = 0
                        optionsDict[ticker][OPTIONS_DATA][current_strike] = strike
                        right = ''
                        if (side == "BUY"):
                            right = 'C'
                            if (float(strike) < valueDict[ticker][LAST]):
                                continue
                        else:
                            right = 'P'
                            if (float(strike) > valueDict[ticker][LAST]):
                                continue  
                        exchange = getExchange(ticker)
                        opt_contract = createContract(ticker, "OPT", exchange, exchange, "USD", right, expiration, strike)
                        print("strike: ", opt_contract.strike)
                        requestOptionsData(ticker, opt_contract)
                        d = optionsDict[ticker][OPTIONS_DATA][delta]
                        print("delta: ", d)
                        if (d == 0):
                            continue
                        elif (d < .50 and d > .30):
                            print("ask price: ", optionsDict[ticker][OPTIONS_DATA][ask])
                            print("max risk: ", max_option_cost)
                            if (optionsDict[ticker][OPTIONS_DATA][ask] * 100 > max_option_cost):
                                print("options are too expensive")
                                continue
                            else:
                                found_option = True
                                print("found option: ", exp, strike)
                                print(optionsDict[ticker][OPTIONS_DATA])
                                break
                        elif (d < .3):
                            break
                        else:
                            continue  
        except TypeError:
            print("no options chain") 
        if (not found_option):
            print("no option was found")
            send(ticker, shares, side, unit, index, ticker_type, "STK", contract)
        else:
            returned_bid = optionsDict[ticker][OPTIONS_DATA][bid]
            returned_ask = optionsDict[ticker][OPTIONS_DATA][ask]
            spread_percentage = round(percentChange(returned_ask, returned_bid), 2)
            print("spread", spread_percentage)
            if (spread_percentage > max_spread_percentage):
                print("spread is too wide")
                send(ticker, shares, side, unit, index, ticker_type, "STK", contract)
            else:
                print("sending order")
                #do order stuff
                send(ticker, shares, side, unit, index, ticker_type, "OPT", opt_contract)
                print(optionsDict)
    except Exception as e:
        print("new order error: " + str(e))

def wait():
    while (dataId not in id_list):
        if (dataId in error_list):
            print("error retrieving request")
            break
        sleep(.1)

def requestOptionsData(ticker, contract_info):
    global dataId, request_options
    request_options = True
    tws.reqMktData(dataId, contract_info, "101", False, False, [])
    start = time.time()
    waiting = True
    errors = [0, -1]
    while ((optionsDict[ticker][OPTIONS_DATA][bid] in errors or optionsDict[ticker][OPTIONS_DATA][ask] in errors or optionsDict[ticker][OPTIONS_DATA][delta] in errors or optionsDict[ticker][OPTIONS_DATA][gamma] in errors) and waiting and dataId not in error_list):
        end = time.time()
        if (end - start > timeout):
            print("timeout")
            waiting = False
    request_options = False
    dataId+=1    

#we already have order details for closing, so this works a bit differently
def sendCloseOrder(ticker, status, shares, side, index, ticker_type):
    global order_id, request_options
    unit = getUnit(ticker, index)
    shares = shares
    if (status == "juice"):
        if (ticker not in momoDict):
            shares = int(shares/2)
        else:
            shares = int(shares/1.5)
    #check to see if stock or option
    instrument, right = "OPT", ''
    log_index = int(valueDict[ticker][LOGGINGID]) - log_offset
    print(ticker, ' ', log_index)
    if (expirations_list[log_index] == ''):
        instrument = "STK"
    if (instrument == "STK"):
        if (side == "BUY"):
            price = round(valueDict[ticker][LAST] + (valueDict[ticker][LAST] * .0075), 1)#round to conform to minimum price increment
        else:
            price = round(valueDict[ticker][LAST] - (valueDict[ticker][LAST] * .0075), 1)
        order = createOrder('LMT', shares, side, price)
        tws.placeOrder(order_id, contract, order)
        sleep(.1)
        print("close order placed for %s" % (ticker))
    else:
        expiration = expirations_list[log_index]
        print("expiration for closing ", expiration)
        strike = strikes_list[log_index]
        if (side == "BUY"): #in this case we are buying back a short, so rights will be flipped around
            right = 'P'
        else:
            right = 'C'
        request_options = True
        exchange = getExchange(ticker)
        opt_contract = createContract(ticker, "OPT", exchange, exchange, "USD", right, expiration, strike)
        print(opt_contract)
        requestOptionsData(ticker, opt_contract)
        bid_, ask_ = 0, 1
        bid_, ask_ = optionsDict[ticker][OPTIONS_DATA][bid_], optionsDict[ticker][OPTIONS_DATA][ask_]
        print(optionsDict[ticker][OPTIONS_DATA])
        if (status == "win" or status == "loss"):
            price = round(bid_ - (bid_ * .02), 2)
            order = createOrder("LMT", shares, "SELL", price) #always buying to open in this system
            tws.placeOrder(order_id, opt_contract, order)
            sleep(.1)
            print("close order placed for %s" % (ticker))
        else:
            #difference = abs(ask_ - bid_) #looking to split the spread
            #price = ask_ - (round(difference/2, 2))
            price = ask_
            order = createOrder("LMT", shares, "SELL", price) 
            tws.placeOrder(order_id, opt_contract, order)
            sleep(.1)
            print("close order placed for %s" % (ticker))
    tws.reqExecutions(10001, ExecutionFilter())
    sleep(.1)
    checkExecutions("close", getUnit(ticker, index), order_id)
    for key, value in executionDict.items():
        print(key, value)
    system.update_acell('L' + str(index + offset), order_id)
    close_id_list[index] = order_id
    if (status != "half"):
        logData("close", ticker, "", status, log_index, unit, ticker_type)
    updateStatus(index, status)
    request_options = False
    order_id+=1

#clears execution infomration once index has been closed out
def updateStatus(row, winType):
    status = ""
    if (winType == "juice"):
        status = "juice"
    elif(winType == "win"):
        status = "win"
    elif(winType == "half"):
        status = "half"
    else:
        status = "loss"
    system.update_acell('K' + str(row + offset), status)
    status_list[row] = status

#update sheet with execution details
def printExecutions(ticker, shares, order_id, expiration, strike, right):
    next_available = nextAvailableCell("orders")
    data_to_log = []
    data_to_log.append(ticker)
    execution_list[next_available - offset] = ticker
    data_to_log.append(order_id)
    order_id_list[next_available - offset] = order_id
    data_to_log.append(shares)
    shares_list[next_available - offset] = shares
    data_to_log_range = system.range('G' + str(next_available) + ":I" + str(next_available))
    i = 0
    for cell in data_to_log_range:
        cell.value = data_to_log[i]
        i+=1
    system.update_cells(data_to_log_range)
    if (momoDict[ticker][RISK] != 0):
        system.update_acell('M' + str(next_available), momoDict[ticker][RISK])
        peak_trough_list[next_available - offset] = momoDict[ticker][RISK]
    data_log.update_acell('AH' + str(valueDict[ticker][LOGGINGID]), shares)
    if (optionsDict[ticker][OPTIONS_DATA][delta] != 0): #this will not be 0 if there was an option found
        data_to_log = [] #separate list for options data
        data_to_log.append(expiration)
        expirations_list[int(valueDict[ticker][LOGGINGID]) - log_offset] = expiration
        data_to_log.append(strike)
        strikes_list[int(valueDict[ticker][LOGGINGID]) - log_offset] = strike
        data_to_log.append(right)
        rights_list[int(valueDict[ticker][LOGGINGID]) - log_offset] = right
        data_to_log.append(str(optionsDict[ticker][OPTIONS_DATA]))
        row = str(valueDict[ticker][LOGGINGID])
        data_to_log_range = data_log.range('AC' + row + ":AF" + row)
        i = 0
        for cell in data_to_log_range:
            cell.value = data_to_log[i]
            i+=1
        data_log.update_cells(data_to_log_range)
        #clear options data to ensure update comes though next time there is a request
        optionsDict[ticker][OPTIONS_DATA][bid] = 0
        optionsDict[ticker][OPTIONS_DATA][ask] = 0

#next available cell to place an order
def nextAvailableCell(location):
    cells = ""
    if (location == "orders"):
        cells = system.range("G3:G100")
    else:
        cells = system.range("A3:A100")
    starting_cell = offset
    for cell in cells:
        if (cell.value != ''):
            starting_cell+=1
            continue
        else:
            print("starting cell: %d" % (starting_cell))
            return(starting_cell)

def getExchange(ticker):
    exchange = "SMART"
    if (ticker in nyse_list):
        exchange = "NYSE"
    if (ticker in nasdaq_list):
        exchange = "ISLAND"
    return(exchange)

#Takes care of most of the data that gets logged. Can probably be its own class
def logData(phase, ticker, open_unit, side_status, index, unit, ticker_type):
    global gc, wks, system, data_log
    date = str(datetime.now())
    time = date[:16]
    logged = False
    data_to_log = []
    while (logged == False):
        try:
            if (phase == "open"):
                if (valueDict[ticker][LOGGINGID] == 0):
                    next_available = 2
                    data_log_cells = data_log.range("A2:A1000")
                    for cell in data_log_cells:
                        if (cell.value == ''):
                            break
                        next_available+=1
                    valueDict[ticker][LOGGINGID] = next_available
                    print("logging id set to: ", valueDict[ticker][LOGGINGID])
                    system.update_acell('F' + str(index + offset), valueDict[ticker][LOGGINGID])
                    logging_id_list[index] = valueDict[ticker][LOGGINGID]
                row = str(valueDict[ticker][LOGGINGID])
                data_to_log.append(ticker)
                print("loggging ticker to data log: " + ticker)
                data_to_log.append(str(screenerDict[ticker][CLOUDSCORES]))
                print("logging cloud scores for {}: {}".format(ticker, screenerDict[ticker][CLOUDSCORES]))
                skips = ['C', 'D', 'E', 'F']
                for skip in skips:
                    data_to_log.append("skip")
                # if (open_unit == "one"):
                data_to_log.append(time)
                data_to_log.append("skip")
                # else:
                #     data_to_log.append('H' + row, time)
                data_to_log.append(side_status)
                skips = ['J', 'K', 'L', 'M', 'N', 'O']
                for skip in skips:
                    data_to_log.append("skip")
                #scrape industry
                try:
                    page = requests.get("https://finance.yahoo.com/quote/" + ticker + "/profile?p=" + ticker)
                    soup = BeautifulSoup(page.content, 'html.parser')
                    item = soup.find("p", {"class":"D(ib) Va(t)"})
                    string = str(item.text)
                    substring1 = "Industry:"
                    substring2 = "Full"
                    result = string[(string.index(substring1)+len(substring1)):string.index(substring2)]
                    data_to_log.append(str(result))
                except Exception as e:
                    print("error scraping industry")
                    data_to_log.append("skip")
                #scrape market cap
                try:
                    page = requests.get("https://finance.yahoo.com/quote/" + ticker + "?p=" + ticker)
                    soup = BeautifulSoup(page.content, 'html.parser')
                    item = soup.find("td", text = "Market Cap")
                    market_cap = item.find_next_sibling("td").text
                    data_to_log.append(market_cap)
                except Exception as e:
                    print("error scraping market cap")
                    data_to_log.append("skip")
                data_to_log.append("skip")
                opening_percentage = screenerDict[ticker][PERCENTCHANGE]
                if (screenerDict[ticker][OPEN] < screenerDict[ticker][CLOSE]): #if opening gap is negative
                    opening_percentage*=-1
                data_to_log.append(opening_percentage)
                if (ticker_type == "fade" and momoDict[ticker][PEAK_TROUGH] != 0 and recent_highs_lows_list[index] != -1):
                    ticker_type = "fade cup"
                if (ticker_type == "trend"):
                    ticker_type = "cup"
                data_to_log.append(ticker_type)
                data_to_log.append(valueDict[ticker][PRICE2BOOK])
                data_to_log.append(valueDict[ticker][PRICEEARNINGS])
                if (ticker in screenerDict and float(screenerDict[ticker][EARNINGS]) != 0):
                    data_to_log.append(screenerDict[ticker][EARNINGS])
                else:
                    data_to_log.append("skip")
                data_to_log.append(scrapeVix())
                try:
                    trailing_pe, forward_pe = 0, 1
                    data_to_log.append(getPERatios(ticker)[trailing_pe])
                    data_to_log.append(getPERatios(ticker)[forward_pe])
                except Exception as e:
                    print("no forward earnings available")
                    data_to_log.append("skip")
                    data_to_log.append("skip")
                data_to_log.append("skip")
                data_to_log.append(checkTradingVolumeToday(ticker))
                skips = ['AC', 'AD', 'AE', 'AF', 'AG', 'AH']
                for skip in skips:
                    data_to_log.append("skip")
                data_to_log.append(account_size)  
                data_to_log.append("skip")                       
                data_to_log.append(momoDict[ticker][RISK])
                data_to_log.append("skip")
                data_to_log.append(screenerDict[ticker][VOLUME])
                data_to_log.append("skip")
                try:
                    page = requests.get("https://finance.yahoo.com/quote/SPY/")
                    soup = BeautifulSoup(page.content, 'html.parser')
                    last_price = soup.find("span", {"class":"Trsdu(0.3s) Fw(b) Fz(36px) Mb(-4px) D(ib)"})
                    last_price = float(last_price.text)
                    print("last price: {}".format(last_price))
                    open_price = soup.findAll("td", {"class":"Ta(end) Fw(600) Lh(14px)"})[1]
                    open_price = float(open_price.text)
                    print("open price: {}".format(open_price))
                    previous_close_price = soup.findAll("td", {"class":"Ta(end) Fw(600) Lh(14px)"})[0]
                    previous_close_price = float(previous_close_price.text)
                    print("previous close price: {}".format(previous_close_price))
                    open_to_last_percentage = percentChange(open_price, last_price)
                    previous_close_to_last_percentage = percentChange(previous_close_price, last_price)
                    if (open_price > last_price):
                        open_to_last_percentage*=-1
                    if (previous_close_price > last_price):
                        previous_close_to_last_percentage*=-1
                    data_to_log.append(str([open_to_last_percentage, previous_close_to_last_percentage]))
                except Exception as e:
                    print("spy data error", e)
                    data_to_log.append("skip")
                entry_percentage_from_open = percentChange(screenerDict[ticker][OPEN], valueDict[ticker][LAST])
                if (valueDict[ticker][LAST] < valueDict[ticker][OPEN]):
                    entry_percentage_from_open*=-1
                data_to_log.append(entry_percentage_from_open)
                try:
                    page = requests.get("https://finance.yahoo.com/quote/" + ticker + "/key-statistics?p=" + ticker)
                    soup = BeautifulSoup(page.content, 'html.parser')
                    float_ = soup.findAll("td", {"class":"Fw(500) Ta(end) Pstart(10px) Miw(60px)"})[10]
                    float_ = float_.text
                    shares_short_percentage_of_float = soup.findAll("td", {"class":"Fw(500) Ta(end) Pstart(10px) Miw(60px)"})[15]
                    shares_short_percentage_of_float = shares_short_percentage_of_float.text
                    print("float: {} shares short percent of float: {}".format(float_, shares_short_percentage_of_float))
                    data_to_log.append(str([float_, shares_short_percentage_of_float]))
                except Exception as e:
                    print("float stats error: ", e)
                    data_to_log.append("skip")
                data_to_log.append("skip") #AR column
                print(data_to_log)
                data_to_log_range = data_log.range('A' + row + ":AR" + row)
                i = 0
                for cell in data_to_log_range:
                    if (data_to_log[i] == "skip"):
                        i+=1
                        continue
                    cell.value = data_to_log[i]
                    i+=1
                data_log.update_cells(data_to_log_range)
                logged = True
            else: #close phase
                try:
                    #calculate percentage win and loss/update log
                    close_percentage = round(percentChange(float(data_log.acell('C' + str(index + log_offset)).value), close_one_list[index]), 2)
                    if (side_status == "win" or side_status == "juice"):
                        data_log.update_acell('L' + str(index + log_offset), close_percentage) 
                        logged = True
                        data_win_one_list[index] = close_percentage
                    else:
                        data_log.update_acell('N' + str(index + log_offset), close_percentage)
                        logged = True
                        data_loss_one_list[index] = close_percentage
                    data_log.update_acell('J' + str(index + log_offset), time)
                except Exception as e:
                    print("closing price has not yet been filled: " + str(e))
                    break
        except (Exception) as e:
            if (str(e) == "\'NoneType\' object has no attribute \'text\'"):
                break
            print("data logging error: " + str(e))
            reConnect()
            continue

#it looks like a 0 shows up in the data log when logging close orders if the fill comes in late
def reCheckClosingPercentages():
    i = 0
    closed_statuses = ["juice", "win", "loss"]
    for cell in execution_list:
        status = status_list[i]
        if (status in closed_statuses):
            unit = getUnit(cell, i)
            logging_id = 0
            ii = 0
            #get logging id
            for ticker in data_list:
                if (cell == ticker):
                    logging_id = logging_id_list[ii]
                    break
                ii+=1
            logData("close", cell, "", status, int((logging_id - log_offset)), unit, "")
        i+=1

#return vix value for logging/directional bias
def scrapeVix():
    page = requests.get("https://finance.yahoo.com/quote/%5EVIX?p=^VIX")
    soup = BeautifulSoup(page.content, 'html.parser')
    item = soup.find("div", {"class":"My(6px) Pos(r) smartphone_Mt(6px)"})
    response_string = str(item.text)
    vix = int(response_string[0:2])
    return(vix)

#trailing PE and forward PE from yahoo for logging
def getPERatios(ticker):
    page = requests.get("https://finance.yahoo.com/quote/" + ticker +"/key-statistics?p=" + ticker)
    soup = BeautifulSoup(page.content, 'html.parser') 
    def find_item(string):
        item = soup.find(string=string)
        tr = item.find_parent("tr")
        tds = tr.find_all("td")
        i = 1
        for td in tds:
            if (i == 2):
                item = td.text
            i+=1
        return(item)
    trailing_pe, forward_pe = find_item("Trailing P/E"), find_item("Forward P/E")
    return([trailing_pe, forward_pe])

#upon entry, check % of average volume that has traded
def checkTradingVolumeToday(ticker):
    try:
        page = requests.get("https://finance.yahoo.com/quote/" + ticker +"?p=" + ticker)
        soup = BeautifulSoup(page.content, 'html.parser')
        today_volume = float(soup.find("td", {"class":"Ta(end) Fw(600) Lh(14px)", "data-test" : "TD_VOLUME-value"}).text.replace(',', ''))
        average_volume = float(soup.find("td", {"class":"Ta(end) Fw(600) Lh(14px)", "data-test" : "AVERAGE_VOLUME_3MONTH-value"}).text.replace(',', ''))
        percentage_of_average_float_traded = round((today_volume/average_volume) * 100, 2)
        return(percentage_of_average_float_traded)
    except Exception as e:
        print("check trading volume today error ", e)
        return(-1)

#Create an Order object (Market/Limit) to go long/short.
def createOrder(order_type, quantity, action, price):
    order = Order()
    order.orderType = order_type
    order.totalQuantity = quantity
    order.action = action
    order.lmtPrice = price
    return order

#Create a Contract object defining what will be purchased, at which exchange and in which currency.
def createContract(symbol, sec_type, exch, prim_exch, curr, right, exp, strike):
    contract = Contract()
    if (sec_type == "STK"):
        contract.symbol = symbol
        contract.secType = sec_type
        contract.exchange = exch
        contract.primaryExch = prim_exch
        contract.currency = curr
    else:
        contract.symbol = symbol
        contract.secType = sec_type
        contract.exchange = exch
        contract.primaryExch = prim_exch
        contract.currency = curr
        contract.lastTradeDateOrContractMonth = exp
        contract.strike = strike
        contract.right = right
        contract.multiplier = 100
    return(contract)


#returns absolute value  
def percentChange(closePrice, openPrice):
    if (openPrice == 0 or closePrice == 0):
        return(0)
    if (openPrice > closePrice):
        changePercentage = ((openPrice - closePrice)/closePrice)*100
    else: 
        changePercentage = ((closePrice - openPrice)/closePrice)*100
    return(round(changePercentage, 2))

#has an order been placed already? This function isn't really neeeded now that only one unit is used
def canOrder(ticker, ticker_type):
    if (ticker in execution_list):
        return False
    else:
        return True
    
def getUnit(ticker, index): #all trades are only using one unit as of now. Will change with bigger positions to get filled fully
    count = 0
    x = 0
    for cell in execution_list:
        if (cell == ticker):
            temp_index = x
            count+=1
            if (temp_index == int(index) and count == 1):
                return("one")
            else:
                return("two")
        x+=1

#return true if full unit is there
def numUnits(ticker):
    units = 0
    for cell in execution_list:
        if (cell == ticker):
            unit+=1
    if (units == 2):
        return True
    else:
        return False

#updates highs and lows next to data column in system sheet 
def updateRecentHighLow(ticker, side, recent_high_low, index):
    if ((trade_type_list[index] == 'T' or trade_type_list[index] == 'F' or trade_type_list[index] == 'M') and (ticker not in screenerDict or not canOrder(ticker, "trend"))):
        return
    def update_time(index):
        if (trade_type_list[index] == 'T' or trade_type_list[index] == 'F' or trade_type_list[index] == 'M'):
            time = str(datetime.now())
            system.update_acell('B' + str(index + offset), time)
            time_volume_list[index] = time
    if (side == "BOT"):
        if (valueDict[ticker][LAST] > recent_high_low):
            system.update_acell('C' + str(index + offset), valueDict[ticker][LAST])
            recent_highs_lows_list[index] = valueDict[ticker][LAST]
            update_time(index)
            if (ticker in momoDict):
                momoDict[ticker][CHECKING] = True
    else:
        if (recent_high_low == 0):
            system.update_acell('C' + str(index + offset), valueDict[ticker][LAST])
            recent_highs_lows_list[index] = valueDict[ticker][LAST]
            update_time(index)
            if (ticker in momoDict):
                momoDict[ticker][CHECKING] = True
        if (valueDict[ticker][LAST] != 0 and valueDict[ticker][LAST] < recent_high_low):
            system.update_acell('C' + str(index + offset), valueDict[ticker][LAST])
            recent_highs_lows_list[index] = valueDict[ticker][LAST]
            update_time(index)
            if (ticker in momoDict):
                momoDict[ticker][CHECKING] = True
         
#return market cap string
def marketCap(mktcp):
    if (mktcp < 2000):
        return "micro"
    elif (mktcp >= 2000 and mktcp < 10000):
        return "small"
    elif (mktcp >= 10000 and mktcp < 49900):
        return "mid"
    elif (mktcp >= 49900 and mktcp <= 100000):
        return "large"
    else:
        return "mega"   

def reConnect(): #typically used when exception is thrown due to server timeouts with google sheets
    print("reconnecting")
    global gc, wks, system, data_log, reset_count, scope, credentials
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
    gc = gspread.authorize(credentials)
    wks = gc.open("RCP90")
    system = wks.get_worksheet(which_sheet)
    data_log = wks.get_worksheet(which_log)
    system.update_acell("O2", reset_count)
    reset_count+=1
    # tws.disconnect()
    # sleep(.1)
    # tws.connect()
    # sleep(.1)

def getHighestId():
    global order_id
    starting_id = 0
    for cell in order_id_list:
        if (cell > starting_id):
            starting_id = cell
    for cell in close_id_list:
        if (cell > starting_id):
            starting_id = cell
    if (starting_id > order_id):
        order_id = int(starting_id + 1)
        
def findNumberOfDays(date_one, date_two):
    format_ = "%Y%m%d"
    start = datetime.strptime(date_one, format_)
    end = datetime.strptime(date_two, format_)
    delta = start - end
    days = abs(delta.days)
    return(days)

#only shorting if vix is above 15, only longing if vix is below 20, both in between.
def checkVix():
    global longs_off, shorts_off
    longs_off, shorts_off = False, False
    vix_value = scrapeVix()
    print("vix value: " + str(vix_value))
    if (vix_value > 20):
        print("longs off")
        longs_off = True
    if (vix_value <= 15):
        print("shorts off")
        shorts_off = True

#parses string representation of cell object to find column           
def getColNum(cellInfo):
    col = ''
    string = str(cellInfo)
    m = re.search("(\d+).*?(\d+)", string)
    if (m):
        col = m.group(2)
        return(col)

#determines if list values are numbers or strings and populated accordingly 
def populateList(cell, list_, sheet):
    nums_list = []
    if (sheet == "system"):
        nums_list = ['3', '6', '8', '9', '10', '12', '13']
    if (sheet == "log"):
        nums_list = ['3', '4', '5', '6', '30']
    if (sheet == "screen"):
        nums_list = ['5']
    if (getColNum(cell) in nums_list):
        if (cell.value == ''):
            list_.append(0)
        else:
            list_.append(float(cell.value))
    else:
        if (cell.value == ''):
            list_.append('')
        else:
            list_.append(cell.value)

#returns true if there is buying power available
def maxOrders(ticker_type):
    if (available_funds[available] == 0 and available_funds[sma] == 0):
        print("SMA/Available Funds not retrieved")
        return False
    remaining = min(available_funds[available], available_funds[sma]) * 2
    #print("remaining funds: %.2f" % (remaining))
    if (remaining >= 2000):
        return True
    else:
        print("no day funds remaining")
        return False

def findRanges():
    ranges = []
    max_ = 60
    #find number of tickers in screen
    i = 1
    for ticker in screen_tickers_list:
        if (ticker == ''):
            break
        i+=1
    print("range end: ", i)
    range_end = i
    split = float(i/max_)
    whole_nums = int(split)
    remainder = (split - whole_nums) * 100
    last_num = 0
    if (whole_nums > 0):
        for j in range(whole_nums):
            if (j == 0):
                last_num = max_
                ranges.append(['2', str(last_num)])
            else:
                next_last_num = last_num + max_
                ranges.append([str(last_num+1), str(next_last_num)])
                last_num = next_last_num
        if (remainder > 0):
            last_range_start = last_num + 1
            last_range_end = range_end
            ranges.append([str(last_range_start), str(last_range_end)])
    else:
        ranges.append(['2', str(range_end)])
    return(ranges)

#load cells into lists with cell values
print("loading sheet")
data_cells = system.range("A3:A100")
data_list = []
for cell in data_cells:
    populateList(cell, data_list, "system")
time_volume_cells = system.range("B3:B100")
time_volume_list = [] 
for cell in time_volume_cells:
    populateList(cell, time_volume_list, "system")   
long_short_cells = system.range("E3:E100")
long_short_list = []
for cell in long_short_cells:
    populateList(cell, long_short_list, "system")
logging_id_cells = system.range("F3:F100")
logging_id_list = []
for cell in logging_id_cells:
    populateList(cell, logging_id_list, "system")
execution_tickers = system.range("G3:G100")
execution_list = []
for cell in execution_tickers:
    populateList(cell, execution_list, "system")
order_id_cells = system.range("H3:H100")
order_id_list = []
for cell in order_id_cells:
    populateList(cell, order_id_list, "system")
shares = system.range("I3:I100")
shares_list = []
for cell in shares:
     populateList(cell, shares_list, "system")
fill_cells = system.range("J3:J100")
fill_list = []
for cell in fill_cells:
    populateList(cell, fill_list, "system")
status_cells = system.range("K3:K100")
status_list = []
for cell in status_cells:
    populateList(cell, status_list, "system")
close_id_cells = system.range("L3:L100")
close_id_list = []
for cell in close_id_cells:
    populateList(cell, close_id_list, "system")
peak_trough_list = []
peak_trough_cells = system.range("M3:M100")
for cell in peak_trough_cells:
    populateList(cell, peak_trough_list, "system")
fade_tickers = screener.range("D2:D100")
fade_list = []
for cell in fade_tickers:
    populateList(cell, fade_list, "screen")
trend_range_list = []
trend_range_cells = screener.range("F2:F100")
for cell in trend_range_cells:
    populateList(cell, trend_range_list, "screen")
pre_marklet_high_low_cells = screener.range("E2:E100")
pre_market_high_low_list = []
for cell in pre_marklet_high_low_cells:
    populateList(cell, pre_market_high_low_list, "screen")
earnings_surprise = screener.range("H2:H100")
earnings_list = []
for cell in earnings_surprise:
    populateList(cell, earnings_list, "screen")
screen_tickers_list = []
screen_tickers = screener.range("D2:D1000")
for cell in screen_tickers:
    populateList(cell, screen_tickers_list, "screen")
premarket_volume_list = []
premarket_volume_cells = screener.range("I2:I1000")
for cell in premarket_volume_cells:
    populateList(cell, premarket_volume_list, "log")
close_one_cells = data_log.range("E2:E1000")
close_one_list = []             
for cell in close_one_cells:
    populateList(cell, close_one_list, "log")
close_two_cells = data_log.range("F2:F1000")
close_two_list = []             
for cell in close_two_cells:
    populateList(cell, close_two_list, "log")
data_win_one_cells = data_log.range("L2:L1000")
data_win_one_list = []             
for cell in data_win_one_cells:
    populateList(cell, data_win_one_list, "log")
data_win_two_cells = data_log.range("M2:M1000")
data_win_two_list = []             
for cell in data_win_two_cells:
    populateList(cell, data_win_two_list, "log")
data_loss_one_cells = data_log.range("N2:N1000")
data_loss_one_list = []             
for cell in data_loss_one_cells:
    populateList(cell, data_loss_one_list, "log")
data_loss_two_cells = data_log.range("O2:O1000")
data_loss_two_list = []             
for cell in data_loss_two_cells:
    populateList(cell, data_loss_two_list, "log")
exit_one_list = []
exit_one_cells = data_log.range("J2:J1000")
for cell in exit_one_cells:
    populateList(cell, exit_one_list, "log")
exit_two_list = []
exit_two_cells = data_log.range("K2:K1000")
for cell in exit_two_cells:
    populateList(cell, exit_two_list, "log")
expirations_list = []
expiration_cells  = data_log.range("AC2:AC1000")
for cell in expiration_cells:
    populateList(cell, expirations_list, "log")
strikes_list = []
strike_cells  = data_log.range("AD2:AD1000")
for cell in strike_cells:
    populateList(cell, strikes_list, "log")
rights_list = []
right_cells  = data_log.range("AE2:AE1000")
for cell in right_cells:
    populateList(cell, rights_list, "log")

date = str(datetime.now())
if (date[11:16] == "09:29"):
    subprocess.call(["afplay", "Wunderwaffe.mp3"])

fades_on = False
momos_on  = False
cups_on = False
current_time = datetime.strptime(str(datetime.now())[11:19].replace('-', ''), "%H:%M:%S") 
fade_start_time = datetime.strptime("09:34:00", "%H:%M:%S")
momo_start_time = datetime.strptime("09:50:00", "%H:%M:%S")
cup_start_time = datetime.strptime("09:50:00", "%H:%M:%S")
fade_difference = current_time - fade_start_time
momo_difference = current_time - momo_start_time
cup_difference = current_time - cup_start_time
seconds_until_executions_off = 200
try:
    fade_difference_amount = int(str(fade_difference)[0:4].replace(':', ''))
    if (not fade_difference_amount > seconds_until_executions_off):
        if (str(fade_difference)[0] != '-'):
            fades_on = True       
except:
    print("fades off")
try:
    momo_difference_amount = int(str(momo_difference)[0:4].replace(':', ''))
    if (not momo_difference_amount > seconds_until_executions_off):
        if (str(momo_difference)[0] != '-'):
            momos_on = True       
except:
    print("momos off")
try:
    cup_difference_amount = int(str(cup_difference)[0:4].replace(':', ''))
    if (not cup_difference_amount > seconds_until_executions_off):
        if (str(cup_difference)[0] != '-'):
            cups_on = True 
except:
    print("cups off") 


#load screenerDict/apply hack. ranges will vary depending on number of names being screened
#I did this so I could re call the same script multiple times with different slices of data (lists of tickers in this case) without violating data pacing restrictions
#May have been having issues with calling disconnect and then reconnecting and still breaking data rule. Honestly not sure, but this hack works fine for now.
file = open("screener.csv", mode = 'r')
contents = file.read()
string = str(contents)
if (string == ''): 
    print("populating screener")
    ranges = findRanges()
    print(ranges)
    for range_ in ranges:
        sys.argv = ['opening_percentages.py', range_[0], range_[1], socket_port]
        print("starting opening percentages")
        exec(open('opening_percentages.py').read())
    file = open("screener.csv", mode = 'r')
    contents = file.read()
    string = str(contents)
    string = string.replace("}{", ', ')
    screenerDict = ast.literal_eval(string)
    file.close()
else:
    string = string.replace("}{", ', ')
    screenerDict = ast.literal_eval(string)
    file.close()

print("connecting to TWS")
types = (message.tickSize, message.tickPrice, message.tickOptionComputation, message.tickGeneric, message.tickString)
tws = ibConnection(port=socket_port, clientId=100)
tws.register(marketData, *types)
tws.register(execDetails, message.execDetails)
tws.register(accountSummary, message.accountSummary)
tws.register(nextValidId, message.nextValidId)
tws.register(historicalData, message.historicalData)
tws.register(historicalDataEnd, message.historicalDataEnd)
tws.register(symbolSamples, message.symbolSamples)
tws.register(securityDefinitionOptionParameter, message.securityDefinitionOptionParameter)
tws.register(errorHandler, message.error)
#tws.registerAll(reply_handler)
tws.connect()
sleep(.1)
#get starting order id
tws.reqIds(1) #parameter is deprecated
sleep(.1)
#order_id = 200
print("order id: ", order_id)
while (account_size == 0):
    tws.reqAccountSummary(9004, "All", "NetLiquidation") #size positions based on current account size to continuously compound
    sleep(.1)
print("current account size: ", account_size)
timeout, max_spread_percentage, max_option_cost, cup_max_risk, equity_max_risk = 2, 20, account_size * .024, account_size * .0085, account_size * .01
getHighestId()
print("starting id: %d" % (order_id))
vix_checked = False
while (not vix_checked):
    try:
        checkVix()
        vix_checked = True
    except:
        print("check vix error")
        continue
while_iteration = 0
while True:
    try:
        trade_type_cells = system.range("D3:D100")#in case we want to use launchpad on a name that is in the screen
        trade_type_list = []
        for cell in trade_type_cells:
            populateList(cell, trade_type_list, "system")
        recent_highs_lows = system.range("C3:C100")
        recent_highs_lows_list = []
        for cell in recent_highs_lows: 
            populateList(cell, recent_highs_lows_list, "system")
        tws.reqIds(1)
        if (system.acell("A2").value == "launchpad running"):
            reConnect()
            data_cells = system.range("A3:A100")
            data_list = []
            for cell in data_cells:
                populateList(cell, data_list, "system")
            time_volume_cells = system.range("B3:B100")
            time_volume_list = [] 
            for cell in time_volume_cells:
                populateList(cell, time_volume_list, "system")
            recent_highs_lows = system.range("C3:C100")
            recent_highs_lows_list = []
            for cell in recent_highs_lows: 
                populateList(cell, recent_highs_lows_list, "system")   
            trade_type_cells = system.range("D3:D100")
            trade_type_list = []
            for cell in trade_type_cells:
                populateList(cell, trade_type_list, "system")
            long_short_cells = system.range("E3:E100")
            long_short_list = []
            for cell in long_short_cells:
                populateList(cell, long_short_list, "system")
            system.update_acell("A2", '')
            screening = True
        while_iteration+=1 #being used as time buffer to avoid pacing violations
        print("iteration: ", while_iteration)
        if (not tws.isConnected()):
            reConnect()
        date = str(datetime.now())
        print(date)
        print("momos on: ", momos_on, "fades on: ", fades_on, "cups on: ", cups_on)
        print("longs off: ", longs_off, "shorts_off: ", shorts_off)
        if (date[11:16] == "09:34" or date[11:16] == "09:35"):
            print("fades_on")
            fades_on = True
        if (date[11:16] == "09:50" or date[11:16] == "09:51"):
            momos_on = True
            cups_on = True
            print("momos on")
        if (date[11:16] == "11:00" or date[11:16] == "11:01"):
            momos_on = False
            fades_on = False
        # if (date[11:13] == "16" or system.acell("N2").value == "exit"):  
        #     print("checking closing fills")
        #     reCheckClosingPercentages()
        #     print("exiting")
        #     tws.disconnect()
        #     sleep(.2)
        #     file = open("screener.csv", mode = 'w')
        #     file.write("")
        #     file.close()
        #     sys.exit(0)
        #screening portion takes care of opening percentages and gets fades ready     
        if (screening):
            index = 0
            for cell in data_list:
                exchange = getExchange(cell)
                if (cell not in screenerDict and recent_highs_lows_list[index] != 0):
                    cloud_score = ""
                    if (trade_type_list[index] == 'L'):
                        cloud_score = "launchpad"
                    print("updating screener dict")
                    print("long short list: ", long_short_list[index])
                    screenerDict.update({cell : [0, 0, recent_highs_lows_list[index], long_short_list[index], exchange, "", cloud_score, 0, 0]})
                index+=1
            try:
                for key, value in screenerDict.items():
                    earnings_sentiment = float(value[EARNINGS])
                    pre_market_volume = value[VOLUME]
                    if (pre_market_volume == ''):
                        pre_market_volume = 0
                    else:
                        pre_market_volume = float(pre_market_volume)
                    if (value[OPEN] == 0 or value[CLOSE] == 0):
                        tickerInPlay = key
                        contract = createContract(key, 'STK', value[EXCHANGE], value[EXCHANGE], 'USD', "", "", "")
                        tws.reqMarketDataType(market_data_type)
                        tws.reqMktData(dataId, contract, "", False, False, [])
                        sleep(.1)
                        dataId+=1
                    if (value[CLOSE] != 0 and value[OPEN] != 0 and value[PERCENTCHANGE] == 0): 
                        openingFluct = percentChange(value[CLOSE], value[OPEN])
                        value[PERCENTCHANGE] = round(openingFluct, 2)
                        if (value[CLOUDSCORES] != "launchpad"):
                            if (openingFluct > fade_parameters[opening_qualifier_one]): #upgrades get 100, downgrades get -100
                                if (value[CLOSE] < value[OPEN] and value[TREND] != "up"):
                                    value[DIRECTION] = 'S'
                                    if ((value[TREND] == "down" and earnings_sentiment < 0) or (value[TREND] == '' and earnings_sentiment < 0 and pre_market_volume < 50)): 
                                        value[TREND] = "fade"
                                    else:
                                        value[TREND] == ''
                                if (value[CLOSE] > value[OPEN] and value[TREND] != "down"):
                                    value[DIRECTION] = 'L'
                                    if ((value[TREND] == "up" and earnings_sentiment > 0) or (value[TREND] == '' and earnings_sentiment > 0)):
                                        value[TREND] = "fade"
                                    else:
                                        value[TREND] = ''
                            if (openingFluct >= momo_percentage_one and openingFluct <= momo_percentage_two):
                                if (value[CLOSE] < value[OPEN] and value[TREND] == "up" and earnings_sentiment >= 0 and value[VOLUME] != ''):
                                    value[DIRECTION] = 'L'
                                    value[TREND] = "trend"
                                if (value[CLOSE] > value[OPEN] and value[TREND] == "down" and earnings_sentiment < 0 and value[VOLUME] != ''):
                                    value[DIRECTION] = 'S'
                                    value[TREND] = "trend"
            except Exception as e:
                print("pre screen error: " + str(e))
            print("open/close dict")
            empty_count = 0
            for key, value in screenerDict.items():
                if (value[OPEN] == 0):
                    empty_count+=1
                if (value[DIRECTION] != '' and value[TREND] != '' and key not in data_list):
                    trade_type = ''
                    if (value[TREND] == "fade"):
                        trade_type = 'F'
                    elif (value[TREND] == "trend"):
                        trade_type = 'T'
                    else:
                        continue
                    next_available = nextAvailableCell("data")
                    data_to_log = []
                    data_to_log.append(key)
                    data_to_log.append("skip")
                    data_to_log.append("skip")
                    data_list[next_available - offset] = key
                    data_to_log.append(trade_type)
                    trade_type_list[next_available - offset] = trade_type
                    data_to_log.append(value[DIRECTION])
                    long_short_list[next_available - offset] = value[DIRECTION]
                    data_to_log_range = system.range('A' + str(next_available) + ":E" + str(next_available))
                    i = 0
                    for cell in data_to_log_range:
                        if (data_to_log[i] == "skip"):
                            i+=1
                            continue
                        cell.value = data_to_log[i]
                        i+=1
                    system.update_cells(data_to_log_range)
            if (empty_count == 0):
                screening = False
        for cell in data_list:
            if (cell == ''):
                break
            if (cell not in valueDict):
                valueDict.update({cell : [0, '', 0, 0, 0, 0, False, 0, 0, 0, 0]})
                optionsDict.update({cell : [0, 0, 0, [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]]})
         #get current logging IDs
        i = 0
        for cell in data_list:
            if (cell == ''):
                break
            elif (logging_id_list[i] == 0):
                i+=1
            else:
                valueDict[cell][LOGGINGID] = logging_id_list[i]
                i+=1
        try:
            #update executionDict/delayed opening fills
            i = 0
            for cell in fill_list:
                if (cell != 0 and order_id_list[i] not in executionDict):
                    executionDict.update({order_id_list[i] : [execution_list[i], long_short_list[i], cell, 0]})
                if (cell == 0 and execution_list[i] != ''):
                    print("checking fill for %s" % (execution_list[i]))
                    temp_id = order_id_list[i]
                    print(temp_id)
                    tws.reqExecutions(10001, ExecutionFilter())
                    sleep(.1)
                    print("second check executions")
                    checkExecutions("open", getUnit(execution_list[i], i), temp_id)
                i+=1
        except Exception as e:
            print("error 1:" + str(e))
        try:
            #update delayed close fills
            i = 0
            for cell in status_list:
                if ((cell == "win" or cell == "loss" or cell == "juice" or cell == "half")):
                    try:
                        unit = getUnit(execution_list[i], i)
                        temp_id = close_id_list[i]
                        index = int(valueDict[execution_list[i]][LOGGINGID]) - log_offset
                        if (close_one_list[index] == 0 or close_two_list[index] == 0):
                            tws.reqExecutions(10001, ExecutionFilter())
                            sleep(.1)
                            checkExecutions("close", unit, close_id_list[i])
                            if (cell != "half"):
                                if (close_one_list[index] != 0 and data_win_one_list[index] == '' and data_loss_one_list[index] == ''):
                                    logData("close", execution_list[i], "", status_list[i], index, unit, '')
                                if (close_two_list[index] != 0 and data_win_two_list[index] == '' and data_loss_two_list[index] == ''):
                                    logData("close", execution_list[i], "", status_list[i], index, unit, '')
                    except Exception as e:
                        print("error 6: " + str(e))
                i+=1
        except Exception as e:
            print("error 3: " + str(e))
        index = 0
        for cell in data_list:
            #tickerInPlay is used as a reference for which stock we are requesting data for/executing with each iteration
            if (cell == ''):
                break
            tickerInPlay = cell
            print("ticker: " + tickerInPlay + ' : ' + str(valueDict[tickerInPlay][LAST]))
            side = ""
            if (long_short_list[index] == 'L'):
                side = "BOT"
            else:
                side = "SLD"
            ticker_type = ""
            if (trade_type_list[index] == 'F'):
                ticker_type = "fade"
            elif (trade_type_list[index] == 'M'):
                ticker_type = "momo"
            elif (trade_type_list[index] == 'L'):
                ticker_type = "launch"
            else:
                ticker_type = "trend"
            if (cell in execution_list):
                i = 0
                for stock in execution_list:
                    if (cell == stock):
                        break
                    i+=1
                if (cell not in momoDict and fill_list[i] != 0 and peak_trough_list[i] != 0):
                    peak = 0
                    if (side == "BOT"):
                        peak = fill_list[i] - (fill_list[i] * (peak_trough_list[i]/100))
                    else:
                        peak = fill_list[i] + (fill_list[i] * (peak_trough_list[i]/100))
                    print(cell, ticker_type)
                    if (ticker_type == "trend" or ticker_type == "fade"):
                        momoDict.update({cell : [peak_trough_list[i], peak_trough_list[i]*risk_reward_parameters[cup_reward], peak_trough_list[i]*risk_reward_parameters[cup_trail_trigger], peak_trough_list[i]*risk_reward_parameters[cup_trail_execute], peak, False, 0]})
                    else:
                        momoDict.update({cell : [peak_trough_list[i], peak_trough_list[i]*risk_reward_parameters[momo_reward], peak_trough_list[i]*risk_reward_parameters[momo_trail_trigger], peak_trough_list[i]*risk_reward_parameters[momo_trail_execute], peak, False, 0]})
            else:
                momoDict.update({cell : [0, 0, 0, 0, 0, True, 0]})
            print("ticker: %s ticker_type: %s" % (tickerInPlay, ticker_type))
            recent_high_low = recent_highs_lows_list[index]
            if (ticker_type != "launch"):
                if (recent_highs_lows_list[index] == -2):#omit a name entirely
                    index+=1
                    print(tickerInPlay + " no longer being screened")
                    continue
                #trends or momo trades with -1 are omitted
                if (recent_highs_lows_list[index] != -1 and recent_highs_lows_list[index] != -2):
                    updateRecentHighLow(tickerInPlay, side, recent_high_low, index)
                #should not go too far beyond open in either direction before looking for entry
                if (recent_highs_lows_list[index] == -1 and ticker_type != "fade"):
                    index+=1
                    print(tickerInPlay + " no longer being screened")
                    continue
                if (recent_highs_lows_list[index] != 0 and cell not in execution_list and valueDict[tickerInPlay][LAST] != 0 and screenerDict[tickerInPlay][OPEN] != 0):
                    if (momos_on or fades_on):
                        recent_highs_lows_list[index] = float(system.acell('C' + str(index+offset)).value) #in case we manually override trade
                    over_exhaust_percentage = 0
                    t, f = 0, 1
                    if (ticker_type == "trend"):
                        over_exhaust_percentage = over_exhausted[t]
                    else:
                        over_exhaust_percentage = over_exhausted[f]
                    #this currently shuts momo off
                    if (percentChange(screenerDict[tickerInPlay][OPEN], recent_highs_lows_list[index]) > over_exhaust_percentage and (date[11:16] == "09:30" or date[11:16] == "09:31" or date[11:16] == "09:32" or date[11:16] == "09:33") and ticker_type != "momo"):
                        print("omitted - overexhausted")
                        if (ticker_type == "trend"):
                            system.update_acell('D' + str(index + offset), 'M')
                            trade_type_list[index] = 'M'
                            index+=1
                            continue
                        else:
                            system.update_acell('C' + str(index + offset), -1)
                            recent_highs_lows_list[index] = -1
                            index+=1
                            continue
                    if (side == "BOT" and valueDict[tickerInPlay][LAST] < screenerDict[tickerInPlay][CLOSE] and (ticker_type == "trend" or ticker_type == "momo")):
                        system.update_acell('C' + str(index + offset), -1)
                        recent_highs_lows_list[index] = -1
                        print("omitted - dipped below previous close")
                        index+=1
                        continue
                    if (side == "SLD" and valueDict[tickerInPlay][LAST] > screenerDict[tickerInPlay][CLOSE] and (ticker_type == "trend" or ticker_type == "momo")):
                        system.update_acell('C' + str(index + offset), -1)
                        recent_highs_lows_list[index] = -1
                        print("omitted - jumped above previous close")
                        index+=1
                        continue
                    if (side == "BOT" and valueDict[tickerInPlay][LAST] > screenerDict[tickerInPlay][CLOSE] and ticker_type == "fade"):
                        system.update_acell('C' + str(index + offset), -2)
                        recent_highs_lows_list[index] = -2
                        print("omitted - jumped above previous close")
                        index+=1
                        continue
                    if (side == "SLD" and valueDict[tickerInPlay][LAST] < screenerDict[tickerInPlay][CLOSE] and ticker_type == "fade"):
                        system.update_acell('C' + str(index + offset), -2)
                        recent_highs_lows_list[index] = -2
                        print("omitted - dipped below previous close")
                        index+=1
                        continue
            fill_price = fill_list[index]
            trail_status = ""
            if (status_list[index] == "trail"):
                trail_status = "trail"
            else:
                trail_status = ''
            shares = shares_list[index]
            exchange = getExchange(tickerInPlay)
            contract = createContract(tickerInPlay, 'STK', exchange, exchange, 'USD', "", "", "")
            tws.reqMarketDataType(market_data_type)
            if (valueDict[tickerInPlay][MARKETCAP] == '' or valueDict[tickerInPlay][PRICE2BOOK] == 0 or valueDict[tickerInPlay][PRICEEARNINGS] == 0 or valueDict[tickerInPlay][BETA] == 0):
                tws.reqMktData(dataId, contract, "258", False, False, [])
                sleep(.1)
                dataId+=1
            tws.reqMktData(dataId, contract, "", True, False, [])
            sleep(.1)
            dataId+=1
            #check for close order
            if (valueDict[tickerInPlay][LAST] != 0 and tickerInPlay in execution_list):
                try:
                    print("closing")
                    execution_index = 0
                    for cell in execution_list:
                        if (cell == tickerInPlay and fill_list[execution_index] != 0 and (status_list[execution_index] == '' or status_list[execution_index] == "trail" or status_list[execution_index] == "half")):
                            fill_price_cell = float(system.acell('J' + str(execution_index+offset)).value)
                            fill_list[execution_index] = fill_price_cell
                            if (cell in momoDict and momoDict[cell][RISK] != 0):
                                peak_trough_cell = float(system.acell('M' + str(execution_index+offset)).value)
                                peak_trough_list[execution_index] = peak_trough_cell
                                if (ticker_type == "momo" or ticker_type == "launch"):
                                     momoDict[cell][RISK], momoDict[cell][REWARD], momoDict[cell][TRAIL_TRIGGER], momoDict[cell][TRAIL_EXECUTE] = peak_trough_list[execution_index], peak_trough_list[execution_index] * risk_reward_parameters[momo_reward], peak_trough_list[execution_index] * risk_reward_parameters[momo_trail_trigger], peak_trough_list[execution_index] * risk_reward_parameters[momo_trail_execute]
                                else:
                                    momoDict[cell][RISK], momoDict[cell][REWARD], momoDict[cell][TRAIL_TRIGGER], momoDict[cell][TRAIL_EXECUTE] = peak_trough_list[execution_index], peak_trough_list[execution_index] * risk_reward_parameters[cup_reward], peak_trough_list[execution_index] * risk_reward_parameters[cup_trail_trigger], peak_trough_list[execution_index] * risk_reward_parameters[cup_trail_execute]
                            close(tickerInPlay, ticker_type, side, fill_list[execution_index], shares_list[execution_index], status_list[execution_index], execution_index)
                            execution_index+=1
                        else:
                            execution_index+=1
                            continue
                except Exception as e:
                    print("closing error: " + str(e))
                    reConnect()
            #check for new order
            if (date[11:16] in vix_checks and not vix_checks[date[11:16]]):
                checkVix()
                vix_checks[date[11:16]] = True
            if (valueDict[tickerInPlay][LAST] != 0 and canOrder(tickerInPlay, ticker_type) and populated and recent_high_low != 0):
                try:
                    tws.reqAccountSummary(10001, "All", "AvailableFunds, SMA")
                    sleep(.1)
                    if (maxOrders(ticker_type)):
                        if (while_iteration % 5 == 0 or ticker_type == "launch"): #temporary fix to avoid pacing violations
                            print("execution")
                            execution(tickerInPlay, recent_high_low, ticker_type, side, index)
                    else:
                        print("no buying power")
                except Exception as e:
                    print("execution error: " + str(e))
            index+=1
        #show state of data
        print("value dict")
        for key, value in valueDict.items():
            print(key, value)
        print("momo dict")
        for key, value in momoDict.items():
            print(key, value)
        print("open close dict")
        for key, value in screenerDict.items():
            print(key, value)
        print("options dict")
        for key, value in optionsDict.items():
            print(key, value)
        print("execution dict")
        for key, value in executionDict.items():
            print(key, value)
        populated = True     
    except Exception as e:
        print("an error occured: " + str(e))
        sleep(1)
        try:
            reConnect()
        except Exception as e:
            print("another error occured")
            continue
        continue
print("done")
tws.disconnect()
