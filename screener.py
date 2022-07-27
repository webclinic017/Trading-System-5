from ibpythonic import ibConnection, message
from ibapi.contract import Contract
from ibapi.ticktype import TickTypeEnum as tt
#from ibapi.execution import ExecutionFilter
from time import sleep
import time
import json  #3 imports below are for google sheets api
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from bs4 import BeautifulSoup
import requests
from technical_analysis import TechnicalAnalysis
#from signal import signal, SIGPIPE, SIG_DFL
#signal(SIGPIPE, SIG_DFL) 
start = time.time()
#STEP1: screening true, range from 2 to "", volume false, comment out upgrades/downgrades/premarket movers in tickerscrape
#STEP2: run ticker scrape in other sheet, uncommented, and move upgrades/premarket movers over below liquid underlyings
#STEP3: screening false, volume false, and run on range of additional tickers (upgrade/downgrades/premarket movers)
#STEP4: screening false, volume true, run on range of entire array of tickers
#STEP5: run ticker scrape again in other sheet to get final pre market earnings results update
#STEP6: make sure testcsv file for dictionary is empty before starting system
screening_tickers = False
requesting_volume = False
#google sheets api linking
credentials_file = "/Users/jakezimmerman/Documents/python3credentials.json"
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
gc = gspread.authorize(credentials)
wks = gc.open("RCP90")
sys, scr, log = 7, 8, 9
screener = wks.get_worksheet(scr)
system = wks.get_worksheet(sys)
data_log = wks.get_worksheet(log)
headers = ["Earnings Pre Market", "No Pre Market Earnings", "Earnings Post Market", "No Post Market Earnings", "Liquid Underlyings", "Upgrades/Downgrades", 
    "No Upgrades or Downgrades", "Pre Market Movers", "No Pre Market Movers", "News"]

range_start = "198"
if (screening_tickers):
    print("finding tickers with activity")
    exec(open('ticker_scrape_stable.py').read())
    #find end
    screened_ticker_cells = screener.range("D2:D500")
    i = 2
    for cell in screened_ticker_cells:
        print(cell.value)
        if (cell.value == ''): 
            break
        i+=1
    range_end = str(i - 1)
    print("run screen up to row: ", range_end)
else:
    range_end = "199" #this must be filled in if screening_tickers is false
 
#handle message callbacks
def historicalDataHandler(msg):
    global ticker, request_id
    if (msg.reqId == request_id):
        if (not requesting_volume):
            if (not _volume_):
                stocks[ticker][prices].append(msg.bar.close)
                stocks[ticker][highs].append(msg.bar.high)
                stocks[ticker][lows].append(msg.bar.low)
            else:
                average_volume[ticker][vol].append(msg.bar.volume)
                #print(msg.bar.date)
        else:
            stocks[ticker][vol].append(msg.bar.volume)
                    
def historicalDataEnd(msg):
    print(msg.reqId)
    id_list.append(msg.reqId)

def marketDataHandler(msg):
    global ticker
    if (msg.reqId == data_id):
        if (highs_lows[ticker][side] == 'L'):
            if (msg.tickType == tt.HIGH):
                highs_lows[ticker][high_low] = msg.price
                print("high for %s: %.2f" % (ticker, msg.price))
        else:
            if (msg.tickType == tt.LOW):
                highs_lows[ticker][high_low] = msg.price
                print("low for %s: %.2f" % (ticker, msg.price))

def errorHandler(msg):
    global ticker
    #turn feed off for a given stock when certain errors arise
    if (not _volume_):
        if (msg.errorCode == 162 and msg.errorMsg != "No historical data query found for ticker id: " + str(msg.id) and msg.errorMsg != "Historical Market Data Service error message:API historical data query cancelled: " + str(msg.id)):
            stocks[ticker][receiving] = False
            print(msg.errorMsg)
        if (msg.errorCode == 200): #contract ambiguity (so i dont have to restart)
            stocks[ticker][receiving] = False
            print(msg.errorMsg)
        if (msg.errorCode == 366):
            stocks[ticker][receiving] = False
            print(msg.errorMsg)
            pass
    else:
        if (requesting_volume):
            if (msg.errorMsg == "Historical Market Data Service error message:HMDS query returned no data: " + ticker + "@SMART Trades"):
                stocks[ticker][receiving] = False
                print(msg.errorMsg)
            if (msg.errorCode == 162 and msg.errorMsg != "No historical data query found for ticker id: " + str(msg.id) and msg.errorMsg != "Historical Market Data Service error message:API historical data query cancelled: " + str(msg.id)):
                stocks[ticker][receiving] = False
                print(msg.errorMsg)
            if (msg.errorCode == 200): #contract ambiguity (so i dont have to restart)
                stocks[ticker][receiving] = False
                print(msg.errorMsg)
            if (msg.errorCode == 366):
                print(msg.errorMsg)
                pass
        else:
            if (msg.errorCode == 162 and msg.errorMsg != "No historical data query found for ticker id: " + str(msg.id) and msg.errorMsg != "Historical Market Data Service error message:API historical data query cancelled: " + str(msg.id)):
                average_volume[ticker][receiving] = False
                print(msg.errorMsg)
            if (msg.errorCode == 200): #contract ambiguity (so i dont have to restart)
                average_volume[ticker][receiving] = False
                print(msg.errorMsg)
        
    print(msg)

def percentChange(closePrice, openPrice):
    if (openPrice == 0 or closePrice == 0):
        return(0)
    if (openPrice > closePrice):
        changePercentage = ((openPrice - closePrice)/closePrice)*100
    else: 
        changePercentage = ((closePrice - openPrice)/closePrice)*100 * -1
    return(changePercentage)

        
#generates values for ichimoku cloud spans/calculates scores
def ichimokuCloud(horizon, ticker, phase): 
    green_cloud = 0
    red_cloud = 0
    green_count = 0
    red_count = 0
    aggregate_cloud = 0
    aggregate_count = 0
    if (phase == "screener"):
        screen_prices = stocks[ticker][prices]
        screen_highs = stocks[ticker][highs]
        screen_lows = stocks[ticker][lows]
    #step sizes were chosen to keep in line with 130 as index 
    index = 52
    for price in screen_prices[index:]:
        nine_period_high = 0
        nine_period_low = 100000
        twenty_six_period_high = 0
        twenty_six_period_low = 100000
        fifty_two_period_high = 0
        fifty_two_period_low = 100000
        #highs
        i = index - 52
        for high in screen_highs[i:index]:
            if (i >= index - 9 and high > nine_period_high):
                nine_period_high = high
            if (i >= index - 26 and high > twenty_six_period_high):
                twenty_six_period_high = high
            if (high > fifty_two_period_high):
                fifty_two_period_high = high
            i+=1
        #lows
        i = index - 52
        for low in screen_lows[i:index]:
            if (i >= index - 9 and low < nine_period_low):
                nine_period_low = low
            if (i >= index - 26 and low < twenty_six_period_low):
                twenty_six_period_low = low
            if (low < fifty_two_period_low):
                fifty_two_period_low = low
            i+=1
        index+=1
        tenkan_sen = (nine_period_high + nine_period_low)/2
        kijun_sen = (twenty_six_period_high + twenty_six_period_low)/2
        span_a = (tenkan_sen + kijun_sen)/2
        span_b = (fifty_two_period_high + fifty_two_period_low)/2
        #print("span a: %.2f span b: %.2f tenkan: %.2f kijun: %.2f" % (span_a, span_b, tenkan, kijun))
        if (span_a > span_b):
            green_cloud += round(abs(percentChange(span_a, span_b)), 0)
            green_count+=1
        elif (span_b > span_a):
            red_cloud += round(abs(percentChange(span_b, span_a)), 0)
            red_count+=1
        else:
            pass
    print("green cloud: %.2f count: %d red cloud: %.2f count: %d" % (green_cloud, green_count, red_cloud, red_count))
    total_cloud = green_cloud + red_cloud
    total_count = green_count + red_count
    if (phase == "screener"):
        if (total_cloud != 0):
            if (green_cloud > red_cloud):
                aggregate_cloud = round((green_cloud/total_cloud) * 100, 0)
                aggregate_count = round((green_count/total_count) * 100, 0)
            else:
                aggregate_cloud = round((red_cloud/total_cloud) * -100, 0)
                aggregate_count = round((red_count/total_count) * -100, 0)
        if (horizon == "week"):          
            stocks[ticker][scores][week][cloud] = aggregate_cloud
            stocks[ticker][scores][week][count] = aggregate_count
        elif (horizon == "month"):
            stocks[ticker][scores][month][cloud] = aggregate_cloud
            stocks[ticker][scores][month][count] = aggregate_count
        elif (horizon == "three_month"):
            stocks[ticker][scores][three_month][cloud] = aggregate_cloud
            stocks[ticker][scores][three_month][count] = aggregate_count
        elif (horizon == "six_month"):
            stocks[ticker][scores][six_month][cloud] = aggregate_cloud
            stocks[ticker][scores][six_month][count] = aggregate_count
        else:
            stocks[ticker][scores][year][cloud] = aggregate_cloud
            stocks[ticker][scores][year][count] = aggregate_count


def calculateCloud(horizon, ticker, phase):
    getData(horizon, ticker, phase)
    ichimokuCloud(horizon, ticker, phase)

def createContract(symbol, sec_type, exch, prim_exch, curr):
    contract = Contract()
    contract.symbol = symbol
    contract.secType = sec_type
    contract.exchange = exch
    contract.primaryExchange = prim_exch
    contract.currency = curr
    return contract

def getData(horizon, stock, phase):
    global request_id
    if (horizon == "week"):
        bars = "5 mins"
        duration = "2 W"
    elif (horizon == "month"):
        bars = "30 mins"
        duration = "2 M"
    elif (horizon == "three_month"):
        bars = "2 hours"
        duration = "4 M"
    elif (horizon == "six_month"):
          bars = "4 hours"
          duration = "6 M"
    else:
          bars = "1 Day"
          duration = "1 Y"
    populated = False
    while (not populated):
        date, e = "", ""
        if (phase == "screener"):
            temp_date = str(datetime.now())[0:10].replace('-', '')
            date = temp_date + " 16:00:00"
            e = stocks[stock][exch]
            #date = "20200427 09:30:00"
            #print("ticker: %s" % (ticker))
        contract = createContract(stock, "STK", "SMART", e, "USD")
        make_request = False
        if (phase == "screener"):
            if ((len(stocks[stock][prices]) == 0 or len(stocks[stock][highs]) == 0 or len(stocks[stock][lows]) == 0) and stocks[ticker][receiving] == True):
                make_request = True
        if (make_request):
            print("requesting from getData")
            use_regular_trading_hours = 1
            requestHistoricalData(contract, date, duration, bars, use_regular_trading_hours)
        else:
            populated = True
            print("prices retrieved")

def averageVolAgainstMovingAverage(values, length):
    technical_analysis = TechnicalAnalysis(values)
    sma = technical_analysis.sma(values, length)
    volume_list = []
    for value in sma:
        volume_list.append(round(value[0], 2))    
    total_volume = 0
    for value in volume_list:
        total_volume+=value
    average_volume = total_volume/len(volume_list)
    volume_today = volume_list[-1]
    percent_difference = round(percentChange(volume_today, average_volume), 2)
    return(percent_difference)

def preMarketVolume(ticker):
    volume_bars = stocks[ticker][vol]
    if (len(volume_bars) == 0):
        return('')
    prints = 0
    for bar in volume_bars:
        if (bar != 0):
            prints+=1
    if (prints > 5):
        return(prints)
    else:
        return('')

def getExchange(ticker):
    exchange = "SMART"
    if (ticker in nyse_list):
        exchange = "NYSE"
    if (ticker in nasdaq_list):
        exchange = "ISLAND"
    return(exchange)

def requestHistoricalData(contract, request_end_date, duration, bar_size, trading_hours):
    global request_id
    #print("historical data contract: ", contract, request_id)
    conn.reqHistoricalData(request_id, contract, request_end_date, duration, bar_size, "TRADES", trading_hours, 1, False, [])
    received = False
    start_time = time.time()
    while (received == False):
        current_time = time.time()
        if (current_time - start_time > 10):
            conn.cancelHistoricalData(request_id)
            received = True
        if (request_id in id_list):
            print("id in list")
            received = True
    request_id+=1

def connectSheet(): #typically used when exception is thrown due to server timeouts with google sheets
    global gc, wks, system, data_log, screener, scope, credentials, sys, scr, log
    print("reconnecting")
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
    gc = gspread.authorize(credentials)
    wks = gc.open("RCP90")
    screener = wks.get_worksheet(scr)
    system = wks.get_worksheet(sys)
    data_log = wks.get_worksheet(log)
   

request_id = 1
ticker = ""
stocks, average_volume = {}, {}
scores = 0
week, month, three_month, six_month, year = 0, 1, 2, 3, 4
count, cloud = 0, 1
prices, highs, lows, vol, exch, receiving = 1, 2, 3, 4, 5, 6
socket_port = 4000
highs_lows = {}
_volume_ = False
high_low, side = 0, 1
id_list = []
offset = 3
earnings = 1 #switch to 1 for earnings tickers
earningsDict = {"ticker" : ['A', 'D'], "exchange" : ['B', 'E'], "trend" : ['C', 'F'], "cloud" : 'G', "pre_market" : 'I'}
headers = ["Earnings Pre Market", "No Pre Market Earnings", "Earnings Post Market", "No Post Market Earnings", "Pre Market Movers", "No Pre Market Movers", "News", "Database", "Upgrades/Downgrades", "No Upgrades or Downgrades"]
cloud_scores = False
earnings_surprise_index = 10
nyse_list = ["TCS", "GPRO", "FIVE", "KEYS", "T", "GDX", "OIH", "SPCE"]
nasdaq_list = ["INTC", "WING", "BLBD", "MATW", "STIM", "SMTC", "CSCO", "MSFT", "IPDN"]
cloud_score_cells = screener.range(earningsDict["cloud"] + range_start + ':' +  earningsDict["cloud"] + range_end)
cloud_list = []
for cell in cloud_score_cells:
    cloud_list.append('')
trend_cells = screener.range(earningsDict["trend"][earnings] + range_start + ':' +  earningsDict["trend"][earnings] + range_end)
trend_list = []
for cell in trend_cells:
    trend_list.append('')
ticker_cells = screener.range(earningsDict["ticker"][earnings] + range_start + ':' +  earningsDict["ticker"][earnings] + range_end)
data_cells = system.range("A3:A100")
recent_highs_lows = system.range("C3:C100")
side_cells = system.range("E3:E100")
logging_id_cells = system.range("F3:F100")
volume_cells = data_log.range("R2:R1000")
data_log_ticker_cells = data_log.range("A2:A1000")
time_of_entry_cells = data_log.range("G2:G1000")
#load sheets into lists
ticker_list = []
for i in range(int(range_end) - int(range_start) + 1):
    ticker_list.append('')
index = 0
for ticker in ticker_cells:
    exchange = getExchange(ticker.value)
    #stock = ticker.value.replace('.', ' ')
    stock = ticker.value
    if (stock in headers):
        stock = stock + '!'
    stocks.update({stock : [[[0, 0], [0, 0], [0, 0], [0, 0], [0, 0]], [], [], [], [], exchange, True]})
    try:
        ticker_list[index] = ticker.value
    except Exception as e:
        print(e)
    index+=1

volume_list = []
for cell in volume_cells:
    volume_list.append(cell.value)
time_of_entry_list = []
for cell in time_of_entry_cells:
    time_of_entry_list.append(cell.value)
data_log_ticker_list = []
for cell in data_log_ticker_cells:
    data_log_ticker_list.append(cell.value)
logging_id_list = []
for cell in logging_id_cells:
    logging_id_list.append(cell.value)
recent_highs_lows_list = []
for cell in recent_highs_lows:
    recent_highs_lows_list.append(cell.value)
side_list = []
for cell in side_cells:
    side_list.append(cell.value)
data_list = []
index = 0
for cell in data_cells:
    if (cell.value != '' and recent_highs_lows_list[index] == ''):
        highs_lows.update({cell.value : [0, side_list[index]]})
    data_list.append(cell.value)
    index+=1

#connect to api
conn = ibConnection(port=socket_port, clientId=100)
conn.register(historicalDataHandler, message.historicalData)
conn.register(historicalDataEnd, message.historicalDataEnd)
conn.register(marketDataHandler, message.tickPrice)
conn.register(errorHandler, message.error)
conn.connect()
sleep(.1)

if (requesting_volume):
    print("screening volume")
    for stock, value in stocks.items():
        stocks[stock][receiving] = True
    populated = False
    while (not populated):
        for stock, value in stocks.items():
            ticker = stock
            exchange = getExchange(ticker)
            contract = createContract(stock, "STK", exchange, exchange, "USD")
            use_regular_trading_hours = 0 
            request_end_date = str(datetime.now()).replace('-', '')[:17]
            #request_end_date = "20191009 09:30:00"
            duration = "28800 S"
            bar_size = "1 min"
            make_request = False
            if (len(stocks[stock][vol]) == 0 and stocks[ticker][receiving] == True and '!' not in ticker):
                make_request = True
            if (make_request):
                print("requesting data")
                requestHistoricalData(contract, request_end_date, duration, bar_size, use_regular_trading_hours)
            else:
                populated = True
                print("prices retrieved")
    print("refreshing connection to google sheets")
    connectSheet()
    pre_market_volume_cells = screener.range(earningsDict["pre_market"] + range_start + ':' +  earningsDict["pre_market"] + range_end)
    pre_market_volume_list = []
    for cell in pre_market_volume_cells:
        pre_market_volume_list.append('')
    i = 0
    for ticker in ticker_list:
        print("ticker: ", ticker)
        if (ticker in headers):
            i+=1
            continue
        #store pre market volume in list
        pre_market_volume_list[i] = preMarketVolume(ticker)
        i+=1
    i = 0
    for cell in pre_market_volume_cells:
        try:
            cell.value = pre_market_volume_list[i]
            i+=1
        except IndexError:
            break
    screener.update_cells(pre_market_volume_cells)
if (not requesting_volume):
    print("screening trends")
    if (requesting_volume):
        requesting_volume = False
    #arbitrary score   
    horizons = ["week", "month", "three_month", "six_month", "year"]
    bulls_list, bears_list, range_up_list, range_down_list = [], [], [], []
    param_score = 65
    score_difference_max = 40
    index = int(range_start)
    for stock, value in stocks.items():
        ticker = stock
        if ('!' in ticker):
            index+=1
            continue
        for horizon in horizons:
            calculateCloud(horizon, ticker, "screener")
            value[prices].clear()
            value[highs].clear()
            value[lows].clear()
        print("CLOUD ticker: %s week: %.2f month: %.2f 3 month: %.2f, 6 month: %.2f, year: %.2f" % (stock, value[scores][week][cloud], value[scores][month][cloud], value[scores][three_month][cloud], value[scores][six_month][cloud], value[scores][year][cloud]))
        print("COUNT ticker: %s week: %.2f month: %.2f 3 month: %.2f, 6 month: %.2f, year: %.2f" % (stock, value[scores][week][count], value[scores][month][count], value[scores][three_month][count], value[scores][six_month][count], value[scores][year][count]))
        #only week and month are being used currently for this
        if (abs((value[scores][week][cloud] + value[scores][week][count])/2) > param_score and abs((value[scores][month][cloud] + value[scores][month][count])/2) > param_score and abs(value[scores][week][count] - value[scores][week][cloud]) <= score_difference_max):
            if (value[scores][week][cloud] > 0 and value[scores][month][cloud] > 0 and value[scores][three_month][cloud] > 0 and value[scores][week][count] > 0 and value[scores][month][count] > 0 and value[scores][three_month][count] > 0):
                bulls_list.append(stock)
            if (value[scores][week][cloud] < 0 and value[scores][month][cloud] < 0 and value[scores][three_month][cloud] < 0 and value[scores][week][count] < 0 and value[scores][month][count] < 0 and value[scores][three_month][count] < 0):
                bears_list.append(stock)
        if (stock not in bulls_list and stock not in bears_list and (value[scores][week][cloud] + value[scores][week][count])/2 > 0):
            range_up_list.append(stock)
        if (stock not in bulls_list and stock not in bears_list and (value[scores][week][cloud] + value[scores][week][count])/2 < 0):
            range_down_list.append(stock)
        index+=1
    print("bulls list")
    for key in bulls_list:
        print(key)
    print("bears list")
    for key in bears_list:
        print(key)
    print("range up list")
    for key in range_up_list:
        print(key)
    print("range down list")
    for key in range_down_list:
        print(key)

    print("refreshing connection to google sheets")
    connectSheet()
    #print trend direction to sheet
    if (earnings == 1):
        i = 0
        for key, value in stocks.items():
            if (key in bulls_list):
                trend_list[i] = "up"
            if (key in bears_list):
                trend_list[i] = "down"
            else:
                pass
            cloud_list[i] = value[scores]
            i+=1
    i = 0
    for cell in trend_cells:
        cell.value = trend_list[i]
        i+=1
    i = 0
    for cell in cloud_score_cells:
        cell.value = str(cloud_list[i])
        i+=1

    screener.update_cells(trend_cells)
    screener.update_cells(cloud_score_cells)

print("showing stocks dict")
for stock, value in stocks.items():
    print(stock, value)

print("_volume_ on")
_volume_ = True 
index, offset = 0, 2
for cell in volume_list:
    if (cell == ''):
        ticker = data_log_ticker_list[index]
        if (ticker == ''):
            break
        exchange = getExchange(ticker)
        average_volume.update({ticker : [[[0, 0], [0, 0], [0, 0], [0, 0], [0, 0]], [], [], [], [], exchange, True]})
        contract = createContract(ticker, "STK", exchange, exchange, "USD")
        date = time_of_entry_list[index]
        new_date = date[0:10].replace('-', '')
        new_date = new_date + date[10:19]
        use_regular_trading_hours = 1
        duration = "1 Y"
        bar_size = "1 day"
        requestHistoricalData(contract, new_date, duration, bar_size, use_regular_trading_hours)
        values = []
        for value in average_volume[ticker][vol][:-1]:
            values.append([value, 0])
        try:
            data_log.update_acell('R' + str(index + offset), str([averageVolAgainstMovingAverage(values, 50), averageVolAgainstMovingAverage(values, 200)]))
        except ZeroDivisionError as e:
            print("data could not be retrieved for %s" % (cell))
            data_log.update_acell('R' + str(index + offset), "NA")
    index+=1
#score calculated using weighted percentages giving highest weight to 5 day trends
#index = 3 #start at 3 to align with values in google sheet
#check for any pre market volume and indicate if it's there with 'y'
for stock, value in stocks.items():
    stocks[stock][receiving] = True
print("exiting")
end = time.time()
total_time_elapsed = end - start
print("total screening time in seconds: ", total_time_elapsed)
conn.disconnect()
