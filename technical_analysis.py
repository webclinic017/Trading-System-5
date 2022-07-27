#this class is designed to mimic functions available in pine script documentation. 
#pine is a series based language. This implementation considers this feature, and all member methods (indicators) return lists.


class TechnicalAnalysis():
    #prices is a list of lists containing bar data
    def __init__(self, prices):
        self.prices = prices
        self.value, self.date, self.volume, self.high, self.low, self.open = 0, 1, 2, 3, 4, 5 #value = close
        #print(self.rsi(14))
        #print(self.stochastic_rsi(14, 3, 3))
        #print(self.candlestick())

    def mac_d(self, setting):
        pass
    
    #simple moving average
    def sma(self, source, length):
        k = []
        index = length
        for price in source[index-1:]:
            k_sum = 0
            for p in source[index-length:index]:
                k_sum+=p[self.value]
            k.append([k_sum/length, price[self.date]])
            index+=1
        return(k)     

    def stochastic_rsi(self, length, smooth, period_d):
        try:
            rsi_prices = self.rsi(length)                   
            k_values  = []
            look_back = 14
            index = look_back
            for price in rsi_prices[index-1:]:
                highest_high, lowest_low = 0, 150
                for p in rsi_prices[index-look_back:index]:
                    if (p[self.value] > highest_high):
                        highest_high = p[self.value]
                    if (p[self.value] < lowest_low):
                        lowest_low = p[self.value]
                k_values.append([100 * ((price[self.value] - lowest_low)/(highest_high - lowest_low)), price[self.date]])
                index+=1
            percent_k = self.sma(k_values, smooth)  
            percent_d = self.sma(percent_k, period_d)
            return([percent_k, percent_d])
        except Exception as e:
            print("stochastic error: ", e)


    def rsi(self, length):
        gain, loss = [], []
        index = 0
        for price in self.prices[1:]:
            difference = round(price[0] - self.prices[index][self.value], 3)
            if (difference >= 0):
                gain.append([difference, price[self.date]])
                loss.append([0, price[self.date]])
            else:
                gain.append([0, price[self.date]])
                loss.append([difference * -1, price[self.date]])
            index+=1
        def calcAvg(source, length):
            average_list = []
            #calculate first average
            running_total = 0
            for price in source[:length+1]:
                running_total+=price[self.value]
            average = running_total/length
            average_list.append([average, source[length][self.date]])
            #calculate subsequent average
            for price in source[length+1:]:
                average_list.append([((average_list[-1][self.value] * (length-1) + price[self.value])/length), price[self.date]])
            return(average_list)
        average_gain = calcAvg(gain, length)
        average_loss = calcAvg(loss, length)
        #calculate RSI
        rsi = []
        index = 0
        for gain in average_gain:
            if (gain[self.value] == 0):
                rsi.append([0, gain[self.date]])
            elif (average_loss[index][self.value] == 0):
                rsi.append([100, gain[self.date]])
            else:
                rsi.append([100 - (100/(1 + (gain[self.value]/average_loss[index][self.value]))), gain[self.date]])
            index+=1
        #print("RSI")
        #print(rsi)
        return(rsi)

    def volume_profile(self, row_size): #the visual aspect of this is just a histogram
        high, low = 0, 100000
        for price in self.prices:
            if (price[self.high] > high):
                high = price[self.high]
            if (price[self.low] < low):
                low = price[self.low]
        range_ = high - low
        price_increment = round(range_/row_size, 2)
        histogram = {}
        starting, ending = round(low, 2), round(low + price_increment, 2)
        # print("range: ", range_)
        # print("price increment: ", price_increment)
        # print("high: ", high)
        # print("low: ", low)
        # print("row size: ", row_size)
        # print("starting: ", starting)
        for i in range(row_size):
            histogram.update({starting : 0})
            for price in self.prices:
                if (price[self.value] >= starting and price[self.value] < ending):
                    histogram[starting]+=round(price[self.volume], 2)
            starting, ending = ending, ending + price_increment
        #print("displaying histogram dictionary")
        #print(histogram)
        #return price level with most volume
        volume_node_high, point_of_control = 0, 0
        volume_node_second_high, point_of_control_two = 0, 0
        volume_node_third_high, point_of_control_three = 0, 0
        for price, volume in histogram.items():
            if (volume > volume_node_high):
                volume_node_high = volume
                point_of_control = price
        for price, volume in histogram.items():
            if (volume > volume_node_second_high and volume < volume_node_high):
                volume_node_second_high = volume
                point_of_control_two = price
        for price, volume in histogram.items():
            if (volume > volume_node_third_high and volume < volume_node_second_high):
                volume_node_third_high = volume
                point_of_control_three = price
        print("point of control: ", point_of_control)
        print("point of control two: ", point_of_control_two) 
        print("point of control three: ", point_of_control_three)

    def average_true_range(self, look_back):
        atr = []
        #get first atr
        tr_total = 0
        i = 0
        for price in self.prices[:look_back]:
            #print(price)
            one = abs(price[self.high] - price[self.low])
            #print("one: ", one)
            tr = 0
            if (i == 0):
                tr = round(one, 2)
            else:
                two = abs(price[self.high] - self.prices[i - 1][self.value])
                three = abs(price[self.low] - self.prices[i - 1][self.value])
                #print("two: ", two, "three: ", three)
                tr = round(max(one, two, three), 2)
            tr_total+=tr
            i+=1
        first_atr = round(tr_total/look_back, 2)
        atr.append([first_atr, self.prices[i - 1][self.date]])
        #get remaining atrs
        previous_atr = atr[-1][0]
        i = look_back
        for price in self.prices[look_back:]:
            one = abs(price[self.high] - price[self.low])
            two = abs(price[self.high] - self.prices[i - 1][self.value])
            three = abs(price[self.low] - self.prices[i - 1][self.value])
            tr = round(max(one, two, three), 2)
            next_atr = round(((previous_atr * (look_back-1)) + tr)/look_back, 2)
            atr.append([next_atr, price[self.date]])
            previous_atr = next_atr
            i+=1
        #print("displaying atr")
        return(atr)

    def fibonacci_levels(self): #return all desired fib levels
        def get_level(high_to_low, level):
            difference = high - low
            x = difference * level
            y = difference - x
            value = 0
            if (high_to_low):
                value = high - y
            else:
                value = low + y
            return(value)
        high, low = 0, 100000
        high_index, low_index = 0, 0 #need to know which came first to structure fib pull correctly
        high_date, low_date = "", ""
        for index, price in enumerate(self.prices):
            if (price[self.low] < low):
                low = price[self.low]
                low_index = index
                low_date = price[self.date]
            if (price[self.high] > high):
                high = price[self.high]
                high_index = index
                high_date = price[self.date]
        # print("high: %.2f, index: %d, date: %s" % (high, high_index, high_date))
        # print("low: %.2f, index: %d, date: %s" % (low, low_index, low_date))
        #determine direction of fib pull
        high_to_low = True
        if (low_index < high_index): #if low comes first
            high_to_low = False
        levels = {0 : 0, .236 : 0, .382 : 0, .5 : 0, .618 : 0, .65 : 0, .786 : 0, .88 : 0, 1 : 0}
        for level, value in levels.items():
            levels[level] = get_level(high_to_low, level)
        print("displaying fib levels")
        for level, value in levels.items():
            print(level, value)
            

    def exponential_moving_average(self, setting):
        pass

    def find_trend_lines_and_channels(self):
        pass

    def pivot_points(self):
        pass

    #return series of candle descriptions (just red or green for now)
    def candlesticks(self):
        close_, open_, date_ = 0, 5, 1      
        candles = []
        for price in self.prices:
            if (price[close_] > price[open_]):
                candles.append(["green", price[date_]])
            elif (price[close_] < price[open_]):
                candles.append(["red", price[date_]])
            else:
                candles.append(["undecided", price[date_]])
        return(candles)

