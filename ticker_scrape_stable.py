from time import sleep
import time
from functools import reduce
import json  #3 imports below are for google sheets api
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import pandas as pd
import numpy as np
from bs4 import BeautifulSoup
import requests
import re
import copy

first_screening = True
current_time = datetime.strptime(str(datetime.now())[11:19].replace('-', ''), "%H:%M:%S") 
second_screen_start_time = datetime.strptime("08:30:00", "%H:%M:%S")
difference = current_time - second_screen_start_time
try:
    difference_amount = int(str(difference)[0:4].replace(':', ''))
    if (str(difference)[0] != '-'):
        first_screening = False
except:
    print("first screening")
which_sheet = 0
if (first_screening):
    which_sheet = 8
else:
    which_sheet = 10

print("writing to screen: ", first_screening, which_sheet)

#google sheets api linking
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
credentials_file = "/Users/jakezimmerman/Documents/python3credentials.json"
credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
gc = gspread.authorize(credentials)
wks = gc.open("RCP90")
sheet = wks.get_worksheet(which_sheet)

#return list of dictionaries (list one pre market/list two post market)
def getEarningsTickers(): 
    one_day, three_days = 1, 3
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"] #swap sunday and monday when done
    months = {"01" : "January", "02" : "February", "03" : "March", "04" : "April", "05" : "May", "06" : "June",
              "07" : "July", "08" : "August", "09" : "September", "10" : "October", "11" : "November", "12" : "December"}
    day_name, yesterday_name = days[datetime.today().weekday()], days[datetime.today().weekday() - 1]
    today_month, yesterday_month = months[str(datetime.today())[5:7]], months[str(datetime.strftime(datetime.now() - timedelta(one_day), '%Y-%m-%d'))[5:7]] #account for end of month
    day_number, yesterday_number = str(datetime.today())[8:10], str(datetime.strftime(datetime.now() - timedelta(one_day), '%Y-%m-%d'))[8:10]
    today = day_name + ', ' + today_month + ' ' + day_number
    yesterday = yesterday_name + ', ' + yesterday_month + ' ' + yesterday_number
    #need a case to account for the weekend 
    if (day_name == "Monday"):
        yesterday_name, yesterday_number = "Friday", str(datetime.strftime(datetime.now() - timedelta(three_days), '%Y-%m-%d'))[8:10]
        yesterday = yesterday_name + ', ' + yesterday_month + ' ' + yesterday_number
    page = requests.get("https://hosting.briefing.com/cschwab/Calendars/EarningsCalendar5Weeks.htm")
    soup = BeautifulSoup(page.content, 'html.parser')
    def find_tickers(which_day, which_earnings):
        string = soup.find(string=which_day)
        print("which day: ", which_day)
        print("string: ", string)
        table = string.find_parent("table")
        table = table.find_all('tr')
        end_string = "After The Close"
        tickers = {}
        def calculate_earnings_score(which_day):
            try:
                stop, in_post_market = False, False
                for tr in table:
                    if (stop):
                        break
                    tds = tr.find_all('td')
                    #print(tds)
                    i = 0
                    for td in tds:
                        if ((td.text.isupper() and which_day == today) or (td.text.isupper() and which_day == yesterday and in_post_market)):
                            try:
                                num_one, num_two, num_three = 1, 3, 6
                                if (i == 1):
                                    num_one, num_two, num_three = 2, 4, 7
                                    if (tds[i].text != tds[i+1].text):
                                        i+=1
                                        continue
                                actual = float(tds[i+num_one].text)
                                consensus = float(tds[i+num_two].text)
                                year_revs = float(tds[i+num_three].text.replace('%', ''))
                                earnings_surprise = round(percentChange(consensus, actual), 0)
                                print(td.text, actual, consensus, year_revs, earnings_surprise)
                                symbol = td.text.replace('.', '/')
                                if (((earnings_surprise > 2 and earnings_surprise <= 50) and (year_revs >= -2 and year_revs < 15)) or (abs(earnings_surprise) < 2 and year_revs > 2)):
                                    tickers.update({symbol : 50})
                                elif (((earnings_surprise < -2 and earnings_surprise >= -50) and (year_revs <= 2 and year_revs > -15)) or (abs(earnings_surprise) < 2 and year_revs < -2)):
                                    tickers.update({symbol : -50})
                                elif ((earnings_surprise > 50 and year_revs > -2) or (earnings_surprise > -2 and year_revs >= 15)):
                                    tickers.update({symbol : 100})
                                elif ((earnings_surprise < -50 and year_revs < 2) or (earnings_surprise < 2 and year_revs <= -15)):
                                    tickers.update({symbol : -100})
                                else:
                                    tickers.update({symbol : 0})
                            except Exception as e:
                                print("inner calc earnings score error: ", e)
                                tickers.update({td.text : 0})
                        if (td.text == end_string):
                            if (which_day == today):
                                stop = True
                                break
                            else:
                                in_post_market = True
                        i+=1
            except Exception as e:
                print("outer calc earnings error: ", e, i)
        calculate_earnings_score(which_day)
        return(tickers)
    pre_market_tickers, post_market_tickers = {}, {}
    print("today: ", today)
    try:
        pre_market_tickers = find_tickers(today, "pre")
        if (len(pre_market_tickers) == 0):
            print("no earnings from pre market")
        pre_market_tickers = checkTradingVolume(pre_market_tickers)
        addToUniqueTickers(pre_market_tickers)
    except Exception as e:
        print("today earnings error ", e)
    print("yesterday: ", yesterday)
    try:
        post_market_tickers = find_tickers(yesterday, "post")
        if (len(post_market_tickers) == 0):
            print("no earnings from post market")
        post_market_tickers = checkTradingVolume(post_market_tickers)
        addToUniqueTickers(post_market_tickers)
    except Exception as e:
        print("previous day earnings error ", e)
    return([pre_market_tickers, post_market_tickers])

def getUpgradesDowngrades():
    irregular_chars = ['/', '?', '\'', '-', ',', '&']
    upgrades_downgrades = {}
    def update_dictionary(string):
        page = requests.get("https://www.marketbeat.com/ratings/" + string + '/')
        soup = BeautifulSoup(page.content, 'html.parser')
        ticker = soup.find("div", {"class" : "ticker-area"})
        print(ticker)
        table = soup.find("table", {"class":"scroll-table sort-table"})
        for child in table.descendants:
            try:
                text = child.text
                length = len(text)
                if (text.isupper()):
                    skip = False
                    for c in text:
                        if (c in irregular_chars or c.isdigit()):
                            skip = True
                            break
                    if (length > 6):
                        continue
                    if (length % 2 == 0):
                        split = int(length/2)
                        first_half = text[:split]
                        back_half = text[split:]
                        if (first_half == back_half):
                            continue
                    for key, value in upgrades_downgrades.items():
                        if (text.startswith(key)):
                            skip = True
                            break
                    if (not skip):
                        upgrades_downgrades.update({text : string})
            except:
                continue
            #sleep(1)
    update_dictionary("upgrades")
    update_dictionary("downgrades")
    final_dict = checkTradingVolume(upgrades_downgrades)
    temp = copy.deepcopy(final_dict)
    final_dict_no_dups = removeDuplicates(temp)
    for key, value in final_dict.items(): #check to see which database names were upgraded or downgraded
        if (key not in final_dict_no_dups):
            upgrades_downgrades_rechecks.update({key : value})
    return(final_dict_no_dups)

            
def getPreMarketMovers(): 
    tickers = {}
    page = requests.get("http://thestockmarketwatch.com/markets/pre-market/today.aspx")
    soup = BeautifulSoup(page.content, 'html.parser')
    table = soup.find("div", {"id":"tableMovers"})
    links = table.find_all('a')
    for link in links:
        if (link.text.isupper() and link.text not in tickers):
           tickers.update({link.text : 0})
    final_dict =  checkTradingVolume(tickers)
    # tickers = {}
    # page = requests.get("https://www.marketwatch.com/tools/screener/premarket")
    # soup = BeautifulSoup(page.content, 'html.parser')
    # table = soup.find_all("div", {"class":"element element--table overflow--table"})
    # for thing in table:
    #     links = thing.find_all('a')
    #     for link in links:
    #         if (link.text.isupper() and link.text not in tickers):
    #            tickers.update({link.text : 0})
    # final_dict.update(checkTradingVolume(tickers))
    return(removeDuplicates(final_dict))


def removeDuplicates(dictionary):
    duplicates = []
    for ticker, value in dictionary.items():
        if (ticker in unique_tickers):
            duplicates.append(ticker)
    for ticker in duplicates:
        del dictionary[ticker]
    return(dictionary)


def addToUniqueTickers(dictionary):
    for ticker, value in dictionary.items():
        if (ticker not in unique_tickers):
            unique_tickers.append(ticker)
            

#scrape from yahoo finance
def checkTradingVolume(dictionary):
    no_liquidity = []
    minimum_trading_volume = 200000
    for ticker, value in dictionary.items():
        try:
            page = requests.get("https://finance.yahoo.com/quote/" + ticker +"?p=" + ticker)
            soup = BeautifulSoup(page.content, 'html.parser')
            today_volume = float(soup.find("td", {"class":"Ta(end) Fw(600) Lh(14px)", "data-test" : "TD_VOLUME-value"}).text.replace(',', ''))
            average_volume = float(soup.find("td", {"class":"Ta(end) Fw(600) Lh(14px)", "data-test" : "AVERAGE_VOLUME_3MONTH-value"}).text.replace(',', ''))
        except Exception as e:
            print("check trading volume error ", e)
            continue
        if (not (average_volume > minimum_trading_volume)):
            no_liquidity.append(ticker)
    for ticker in no_liquidity:
        #print(ticker + " doesn't have sufficient trading volume")
        del dictionary[ticker]
    return(dictionary)
        

def percentChange(closePrice, openPrice):
    if (openPrice == 0 or closePrice == 0):
        return(0)
    if (openPrice > closePrice):
        changePercentage = ((openPrice - closePrice)/closePrice)*100
        if ((closePrice < 0 and openPrice > 0) or (closePrice < 0 and openPrice < 0)):
            changePercentage = changePercentage * -1
    else:
        changePercentage = ((closePrice - openPrice)/closePrice)*100 * -1
        if (closePrice < 0 and openPrice < 0):
            changePercentage = changePercentage * -1
    return(changePercentage)
    

unique_tickers = []
upgrades_downgrades_rechecks = {}
pre_market, post_market = 0, 1
pre_market_earnings = getEarningsTickers()[pre_market]
post_market_earnings = getEarningsTickers()[post_market]
liquid_underlyings = {}
liquid_underlyings_cells = sheet.range("C2:C235") #include liquid underlyings database
for ticker in liquid_underlyings_cells:
    liquid_underlyings.update({ticker.value : 0})
liquid_underlyings = removeDuplicates(liquid_underlyings)
addToUniqueTickers(liquid_underlyings) 
upgrades_downgrades = getUpgradesDowngrades()
#upgrades_downgrades = {}
addToUniqueTickers(upgrades_downgrades)
#pre_market_movers = {} # in case the website is down
pre_market_movers = getPreMarketMovers()
addToUniqueTickers(pre_market_movers)
headers, offset = 6, 2
range_ = len(unique_tickers) + headers + offset
events = sheet.range("D2:D" + str(range_))
results = sheet.range("H2:H" + str(range_))
events_list, results_list = [], []
for i in range(range_-offset):
    events_list.append('')
for i in range(range_-offset):
    results_list.append('')

if (len(pre_market_earnings) != 0): 
    events_list[0] = "Earnings Pre Market"
else:
    events_list[0] = "No Pre Market Earnings"
i = 1
for ticker, surprise in pre_market_earnings.items():
    events_list[i] = ticker
    results_list[i] = surprise
    i+=1
if (len(post_market_earnings) != 0): 
    events_list[i] = "Earnings Post Market"
else:
    events_list[i] = "No Post Market Earnings"
i+=1
for ticker, surprise in post_market_earnings.items():
    events_list[i] = ticker
    results_list[i] = surprise
    i+=1
upgrades_downgrades_divider = i
events_list[i] = "Database"
i+=1
for underlying in liquid_underlyings:
    events_list[i] = underlying
    results_list[i] = 0
    i+=1
##
if (which_sheet == 10):
    if (len(upgrades_downgrades) != 0): 
        events_list[i] = "Upgrades/Downgrades"
    else:
        events_list[i] = "No Upgrades or Downgrades"
    i+=1
    for ticker, status in upgrades_downgrades.items():
        events_list[i] = ticker
        if (status == "downgrades"):
            results_list[i] = -50 #for now, + or - 3 digit earnings surprise influences trade. i.e. no shorts after an upgrade
        if (status == "upgrades"):
            results_list[i] = 50
        i+=1
    if (len(pre_market_movers) != 0):
        events_list[i] = "Pre Market Movers"
    else:
        events_list[i] = "No Pre Market Movers"
    print("pre market movers")
    i+=1
    for ticker, value in pre_market_movers.items():
        events_list[i] = ticker
        results_list[i] = 0
        i+=1
    events_list[i] = "News"
##
i = 0
for cell in events:
    try:
        cell.value = events_list[i]
        i+=1
    except IndexError:
        print(i)
        break
i = 0
for cell in results:
    try:
        cell.value = results_list[i]
        i+=1
    except IndexError:
        print(i)
        break
sheet.update_cells(events)
sheet.update_cells(results)

#check to see which upgrades/downgrades were in database, and update accordingly
if (which_sheet == 10):
    i = 0
    for event in events:
        if (i <= upgrades_downgrades_divider):
            i+=1
            continue
        ticker = event.value
        for key, value in upgrades_downgrades_rechecks.items():
            if (key == ticker):
                if (value == "downgrades"):
                    results[i].value = -50
                if (value == "upgrades"):
                    results[i].value = 50
        i+=1
    sheet.update_cells(results)
#

