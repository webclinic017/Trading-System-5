# Trading-System

## Description

Day trading system that can identify, execute, log, and manage trades for stocks and stock options. It leverages a third party called ibpythonic (https://github.com/quantrocket-llc/ibpythonic) that works with the Interactive Brokers proprietary api to fetch and manipulate market data, and uses google sheets as an interface and collaborative database. This project was run and managed using real money, it executed thousands of real time trades, and was highly effective in keeping my overall trading framework controlled, organized, and efficient. The system was running for about a year starting in January of 2019, but due to degraded performance over time as a result of changing market conditions, it was turned off. 

## How To Improve

I've personally shifted focus to a longer term investment framework, so I no longer have an interest in moving this project forward. It was a tremendous undertaking and awesome learning experience, but unfortunately lacked enough of an edge to outperform the market sustainably. From a strategy standpoint, it could potentially be improved by introducing more backtesting and paying closer attention to price action/broad market conditions in order to make entries and exists more precise. From a software standpoint, if I were to rewrite it, I'd elect to use pandas as my database instead of dictionaries, and I'd maintain state using csv files locally rather than google sheets.


## Note about files

##### Launchpad

At times, it made sense to override the system and make trade entries manually. There is a process involved with sizing positions, determining whether or not options have enough liquidity to trade, and composing orders. launchpad_app.py and launchpad_test.py were created for submitting orders manually without having to think about the order structuring process. It's a Tkinter based interface that allows a user to simply add a stoploss price and click buy or sell to put together a trade and forward it to the exchange. 

##### Stats

I've included some of the files that were used to determine which data points to use. Trends.py, for example, was used to determine that more success was had betting with trends as opposed to against. In total, there were 15-20 different pieces of logic evaluated prior to making trade entries. Some of these were implemented after making certain observations, while others were implemented as a result of statistical analysis.
