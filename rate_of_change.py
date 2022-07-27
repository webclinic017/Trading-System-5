from time import sleep
import datetime
from datetime import timedelta
import math

#Proprietary algo for identifying cup-like price action with increasing rate of change in the direction of the trade
class rateOfChange():

    def __init__(self, prices, highs, lows, high_low, side, last):
        self.prices = prices
        self.highs = highs
        self.lows = lows
        self.high_low = high_low
        self.side = side
        self.peak_index = 0
        self.start, self.peak = 0, 1
        self.last = last
        print("last price in rate of change: ", self.last)

    def getResult(self):
        print("getting cup rate of change result")
        get_starting_point = self.getStartingPoint(self.prices, self.side, self.high_low)
        if (get_starting_point[self.start] == -1 or get_starting_point[self.peak] == -1):
            return([False, 0])
        else:
            return([self.calcFirstHalf(self.prices, get_starting_point[self.start], self.side), get_starting_point[self.peak]])

    #prices is the list of prices between the recent_high_low and the starting point
    #high_low is beginning of first half of cup, starting_point is the end
    #returns true if first half of prices displays a greater aggregate ROC than the second half
    def calcFirstHalf(self, prices, starting_price, side):
        print("peak index: %.2f" % (self.peak_index))
        if (starting_price == -1):
            print("no starting point")
            return False
        rate_of_change_list = []
        target_range = prices[:self.peak_index+1]
        i = 0
        for price in target_range:
            if (i+1 < len(target_range) - 1):
                rate_of_change_list.append(round(self.rateOfChange(target_range[i+1], price), 2))
            i+=1
        halfway = int(len(rate_of_change_list)/2)
        first_half, back_half = rate_of_change_list[:halfway], rate_of_change_list[halfway:]
        aggregate_roc_first, aggregate_roc_back = 0, 0
        i = 0
        for roc in first_half:
            if (i+1 < len(first_half) - 1):
                aggregate_roc_first+=roc
            i+=1
        i = 0
        for roc in back_half:
            if (i+1 < len(back_half) - 1):
                aggregate_roc_back+=roc
            i+=1
        print("FIRST")
        print(first_half)
        print("BACK")
        print(back_half)
        print("first: %.2f back: %.2f" % (aggregate_roc_first, aggregate_roc_back))
        if (side == "BUY" or side == "BOT"):
            difference = aggregate_roc_back - aggregate_roc_first
            if (aggregate_roc_first < aggregate_roc_back):
                print("concavity level: %.2f" % (difference))
                return True
            else:
                print("concavity level: %.2f" % (difference))
                return False
        else:
            difference = aggregate_roc_first - aggregate_roc_back
            if (aggregate_roc_first > aggregate_roc_back):
                print("concavity level: %.2f" % (difference))
                return True
            else:
                print("concavity level: %.2f" % (difference))
                return False

    #returns absolute value   
    def percentChange(self, closePrice, openPrice):
        if (openPrice == 0 or closePrice == 0):
            return(0)
        if (openPrice > closePrice):
            changePercentage = ((openPrice - closePrice)/closePrice)*100
        else: 
            changePercentage = ((closePrice - openPrice)/closePrice)*100
        return(changePercentage) 

    #the point at which the price is above or below trough/peak by 1/3% of the drop from high/low to trough/peak returns -1 if no appropriate point is found
    def getStartingPoint(self, prices, side, high_low):
        peak_trough_index, up_down_tick = 0, 0
        if (side == "BUY" or side == "BOT"):
            low = 10000
            for price in self.lows:
                if (price < low):
                    low = price
            for price in self.lows:
                if (price == low):
                    break
                peak_trough_index+=1  
            up_down_tick = self.percentChange(high_low, self.lows[peak_trough_index])/3.25
            print("up down: %.2f" % (up_down_tick))
            print("trough: %.2f" % (self.lows[peak_trough_index])) 
        else:
            high = 0
            for price in self.highs:
                if (price > high):
                    high = price
            for price in self.highs:
                if (price == high):
                    break
                peak_trough_index+=1
            up_down_tick = self.percentChange(high_low, self.highs[peak_trough_index])/3.25
            print("up down: %.2f" % (up_down_tick))
            print("peak: %.2f" % (self.highs[peak_trough_index]))
        self.peak_index = peak_trough_index
        print("peak index: %.2f" % (peak_trough_index))
        prices_two = prices[peak_trough_index:]
        if (side == "BUY" or side == "BOT"):
            trough = self.lows[peak_trough_index]
            starting, current_high = 0, 0
            i = peak_trough_index
            for price in prices_two:
                if (current_high == 0):
                    starting = price
                    current_high = price
                if (price > current_high):
                    current_high = price
                if (self.percentChange(starting, current_high) > up_down_tick and self.percentChange(starting, current_high) < (up_down_tick * 2)):
                    self.starting_index = i
                    if (self.percentChange(price, self.last) < .2): #ensure that starting point price isn't too far away from last price
                        print("starting price: %.2f" % (price))
                        print("starting index: %.2f" % (i))
                        return([price, trough])
                i+=1
            return([-1, -1])
        else:
            peak = self.highs[peak_trough_index]
            starting, current_low = 0, 0
            i = peak_trough_index
            for price in prices_two:
                if (current_low == 0):
                    starting = price
                    current_low = price
                if (price < current_low):
                    current_low = price
                if (self.percentChange(starting, current_low) > up_down_tick and self.percentChange(starting, current_low) < (up_down_tick * 2)):
                    self.starting_index = i
                    if (self.percentChange(price, self.last) < .2):
                        print("starting price: %.2f" % (price))
                        print("starting index: %.2f" % (i))
                        return([price, peak])
                i+=1
            return([-1, -1])

    def rateOfChange(self, start, end):
        if (start != 0 and end != 0):
            rate_of_change = float(((start - end)/end) * 100)
        else:
            rate_of_change = 0
        return rate_of_change
