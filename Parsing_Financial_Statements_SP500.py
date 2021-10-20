 #!/usr/bin/python3
import openpyxl
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference, Series
from openpyxl.styles import Font
import requests
import csv
from bs4 import BeautifulSoup
from yahoofinancials import YahooFinancials
import yfinance as yf
import datetime
from datetime import date
from dateutil.relativedelta import relativedelta
import pandas as pd
from yahoo_fin import stock_info as si
from yahoo_earnings_calendar import YahooEarningsCalendar
import textwrap
import math
import numpy as np

import StockAnaPy

################################################################################################3############# VARIABLES ##################################


# Metrics to be taken from the Income Statements.
Accounting_Principles_Income = ["Sales/Revenue", "EPS (Diluted)", "Diluted Shares Outstanding", "Net Income"]

# Metrics to be taken from the Balance Sheets.
Accounting_Principles_Balance = ["Total Current Assets", "Total Current Liabilities", "Net Property, Plant & Equipment", "Cash & Short Term Investments", "Total Shareholders' Equity", "Total Assets", "Total Liabilities", "Short Term Debt", "Long-Term Debt", "Common Equity (Total)", "Intangible Assets"]

# Metrics to be taken from the Cash Flow Statements.
Accounting_Principles_Flow = ["Capital Expenditures", "Free Cash Flow"]

# Combine them all.
Accounting_Principles = Accounting_Principles_Income + Accounting_Principles_Balance + Accounting_Principles_Flow


# All metrics that will be copied onto Excel. To be stored in the Main Pandas DataFrame.
Main_Data = ["Market Cap", "Shares Outstanding", "Dividends", "Next Earnings Date", "Date Income", "Date Balance", "Date Flow", "P/E Ratio", "P/E Ratio 3yravg.", "P/B Ratio", "Current Ratio", "Income/Rev", "Rev/MC", "Book/MC", "CAPEX/Rev", "PP&E/Assets", "Cash/Assets", "Int. Ass/Ass", "Debt/Equity", "EPS at PE15", "EPS at PE20", "EPS Yoy", "EPS Pre+2.Quart", "EPS Pre+1 Quart", "EPS Quart", "EPS 2018", "EPS 2019", "EPS 2020", "Book Value", "PP&E", "Cash & Inv.", "Revenue Yoy", "Rev Pre+2.Quart", "Rev Pre+1 Quart", "Rev Quart", "Rev 2018", "Rev 2019", "Rev 2020", "Income Yoy", "Income Pre+2.Quart", "Income Pre+1 Quart", "Income Quart", "Income 2018", "Income 2019", "Income 2020", "CAPEX", "CAPEX Pre+2.Quart", "CAPEX Pre+1 Quart", "CAPEX Quart", "CAPEX 2018", "CAPEX 2019", "CAPEX 2020", "Free Cash Flow Yoy", "FCF Pre+2.Quart", "FCF Pre+1 Quart", "FCF Quart", "FCF 2018", "FCF 2019", "FCF 2020"]


# Companies that don't seem to work from MarketWatch.
exceptions = ["BRK.B", "AMCR", "AMZN", "BF.B", "ESS", "EVRG", "IR", "RTX", "VNT", "VNO", "LIN", "VTRS", "HWM", "WRK", "OGN", "HIG", "MRNA"]


# Deltas for the price change.
Deltas = [1, 2, 3, 5, 10, 15, 30, 60, 90]


# Combining everything into one list.
combined_columns = ["Company Name", "Price"]
combined_columns = combined_columns + Deltas + Main_Data


Data = 20
Delta_for_div = 365 # How many days back does it download the price data.
Yoy_delta = 4 # 4 quartes per year.
delta_year = 3 # Take only the last three years of data for the metrics.
delta_quarter = 3
i_2 = 0
i_nan = 0
last_year = "2020" # The year before the present year.
file = 'SP_500_Companies' # The file to create all the data.
analysis_file = 'Analysis_SP500' # The file where it's more easy to view all the data in one sheet.

# Set to MMM if starting from the top. Else, input the latest company it downloaded here.
company_to_start_from = "NOW"
col_str = 20
col_end = 67

# TO BE CHANGED TO YOUR FILE PATHS #
file_path = "/Users/jules/" + file + ".xlsx"
wb = load_workbook(file_path)










####################################################################################
####################################################################################



# Get the tickers and put them into the tickers_raw dataframe.
tickers_raw = StockAnaPy.get_SP500_companies(exceptions, i_2)

# Initialize some lists and dataframes.
tickers = pd.DataFrame(columns = range(3), index = range(tickers_raw.shape[0]))
Tickers = []
Industries = []

for i in range(0, tickers_raw.shape[0]) :

    # nan gets put into the list for the some reason so this removes them
    if str(tickers_raw.loc[i, 0]) == 'nan' :
        i_nan = i_nan + 1
        continue

    Tickers.append(tickers_raw.loc[i, 0])
    tick = [tickers_raw.loc[i, 0], tickers_raw.loc[i, 1], tickers_raw.loc[i, 2]]
    tickers.loc[i-i_nan, :] = tick

    print (i, tickers.loc[i-i_nan, 0], tickers.loc[i-i_nan, 1])

# Put all the tickers into the right Sheet corresponding to their Sub Industry.
#StockAnaPy.classify_by_sheet(tickers, 1, file, combined_columns)
#exit()


### Main Matrix where the final data gets put ###
Main = pd.DataFrame(columns = combined_columns, index = Tickers)

# Function that returns the data delta as a string to be thrown into Yahoo Finance
Time_frame = StockAnaPy.Timeframe(Delta_for_div)

# start_i is the ticker number in the ticker list to start at. This is used when the program crashes and needs to be reset at the last downloaded ticker.
start_i = StockAnaPy.start_from_here(tickers, company_to_start_from) - 1

# Function to put all data into sheet 1 of excel in the other file.
#put_into_sheet1(file, 'Companies', analysis_file)
#exit()

# Loop through every ticker and go fetch the data.
for i in Tickers :
    
    print ("The ticker being worked on is:", i, " Number ", i_2)
    
    # Makes sure it starts back at the right ticker and doesn't redownload tickers data.
    if i_2 < start_i :
        i_2 = i_2 + 1
        continue

    # URL's to access all three financial statement pages
    pg = ['https://www.marketwatch.com/investing/stock/IBM/financials/income/quarter',
          'https://www.marketwatch.com/investing/stock/IBM/financials/balance-sheet/quarter', 'https://www.marketwatch.com/investing/stock/IBM/financials/cash-flow/quarter']
    
    # Transform the ticker into a string.
    i = str(i)
    
    # Check if the accounting data has already been uploaded for this quarter
    ind = tickers.loc[i_2, 1]

    # Get historictical Prices for the Ticker
    data = yf.download(i, start=str(Time_frame[1]), end=str(Time_frame[0]))
    print (data)
    
    # Get the Close prices only.
    Low = data['Close']
    Len = len(Low)

    # Get latest price for the Ticker
    try :
        live_price = StockAnaPy.Get_Live_Price_Yahoo(i)
    except AttributeError :
        live_price = Low[-1]
        
    # Upload the latest price to Main
    Main.loc[i, "Price"] = live_price

    # Calculate the delta in % for all the delta thresholds in the Delta list.
    for j in Deltas :

        # % changes uploaded to Main.
        Main.loc[i, j] = (((live_price / Low[Len-j]) - 1))


    ##### - - ACOUNTING PRICINPLES DOWNLOAD - - #####

    # DataFrame for the Annual Metrics.
    Sub_main = pd.DataFrame(columns = range(5), index = Accounting_Principles)
    # DataFrame for the Quarter Metrics
    Sub_main_2 = pd.DataFrame(columns = range(5), index = Accounting_Principles)

    # Perform the same parsing of the Marketwatch website twice, once for the annual metrics and once for the quarter metrics. x = 0 is the quarter ones, x = 1 is the annual ones.
    for x in range(0, 2) :

        # Create the URL's for income, balance sheets, cash flow pages
        if x == 1 :
            for y in range(0, 3) :
                pg[y] = pg[y].replace('quarter', '')
        
            Sub_main = pd.DataFrame(columns = range(5), index = Accounting_Principles)

        pg_financials_income = pg[0].replace('IBM', i)
        pg_financials_balancesheet = pg[1].replace('IBM', i)
        pg_financials_cashflow = pg[2].replace('IBM', i)

        # Get HTML code for both sheets
        soup_income = StockAnaPy.get_html_code(pg_financials_income)
        soup_balancesheet = StockAnaPy.get_html_code(pg_financials_balancesheet)
        soup_cashflow = StockAnaPy.get_html_code(pg_financials_cashflow)
        
        # Make sure the year parsed for every company is the same
        y = StockAnaPy.check_years_table(soup_income, x, last_year)

        Data = []

        for j in Accounting_Principles :
    
            # Look for string in the income, then in the balance sheet
            t = StockAnaPy.which_statement_is_it_in(j, Accounting_Principles_Income, Accounting_Principles_Balance, Accounting_Principles_Flow, soup_income, soup_balancesheet, soup_cashflow)
            if t == "None" :
                continue

            Sub_data = []

            # Get the data, transform into number, put in Sub DataFrame
            values = StockAnaPy.find_data(t)
            numbers = StockAnaPy.transform_to_num(values)
            
            # Add all the numbers to the Sub_Main dataframe. This will allow the script to create the metrics more easily. It is basically copying the MarketWatch table into the Sub_main tables.
            Sub_main.loc[j:j, (5-len(numbers)):5] = numbers

        if x == 0 :
            
            # Transfer the Annuals to the Sub_Main_2
            Sub_main_2.loc[:,:] = Sub_main.loc[:,:]
            
            # Check if the latest column actually has data. Sometimes when the new data is getting added to the website, it is not all there. These functions verify that the latest quarter is usable.
            income = StockAnaPy.complete_column(Sub_main_2, Accounting_Principles_Income)
            balance = StockAnaPy.complete_column(Sub_main_2, Accounting_Principles_Balance)
            flow = StockAnaPy.complete_column(Sub_main_2, Accounting_Principles_Flow)

            # Find all three statement dates. They can differ.
            date_income = StockAnaPy.find_statement_date(soup_income, income)
            date_balance = StockAnaPy.find_statement_date(soup_balancesheet, balance)
            date_flow = StockAnaPy.find_statement_date(soup_cashflow, flow)
            statement_dates = [date_income, date_balance, date_flow]
            
    
    # Put the name in
    Main.loc[i, "Company Name"] = tickers.loc[i_2, 2]

    # Put the statement dates in excel.
    Main.loc[i, "Date Income":"Date Flow"] = [date_income, date_balance, date_flow]
    
    # Append Market Cap to Row
    Market_cap = Sub_main_2.loc["Diluted Shares Outstanding", income] * live_price
    Main.loc[i, "Market Cap"] = Market_cap
    
    # Append Shares Outstanding to Rows. If Shares is False, will get the previous year value.
    Shares = StockAnaPy.CheckLatestShares(Sub_main_2.loc["Diluted Shares Outstanding", income])
        
    if Shares is False :
        Usable_Shares = StockAnaPy.GetLatestShares(Sub_main_2, "Diluted Shares Outstanding", income)
        Main.loc[i, "Shares Outstanding"] = Usable_Shares
            
    else :
        Main.loc[i, "Shares Outstanding"] = Shares

    # Append Market Cap to Row
    Market_cap = Main.loc[i, "Shares Outstanding"] * live_price
    Main.loc[i, "Market Cap"] = Market_cap
    
    # Get dividends history
    ## NEED TO CHANGE THIS TO HAVE IT BE TAKEN FROM YAHOO FINANCE ##
    Main.loc[i, "Dividends"] = 0
    
    # Get next earnings date
    ## NEED TO CHANGE THIS TO HAVE IT BE TAKEN FROM YAHOO FINANCE ##
    Main.loc[i, "Next Earnings Date"] = 0
    
    # Get P/E ratio, EPS
    eps = StockAnaPy.adding_data_past_data(Sub_main, Sub_main_2, "EPS (Diluted)", income, delta_year, delta_quarter, y)
    Main.loc[i, "P/E Ratio"] = live_price/eps[0]
    Main.loc[i, "EPS Yoy":"EPS 2020"] = eps
    
    # Get the P/E Ratio Avg. over the last 3 years.
    PE_Average = pd.Series.mean(Main.loc[i, "EPS 2018":"EPS 2020"])
    Main.loc[i, "P/E Ratio 3yravg."] = live_price/PE_Average
        
    # Get the EPS for a PE of 15
    Main.loc[i, "EPS at PE15"] = live_price/15
        
    # Get the EPS for a PE of 20
    Main.loc[i, "EPS at PE20"] = live_price/20
    
    # Get Book Value, P/B
    book_value = Sub_main_2.loc["Total Assets", balance] - Sub_main_2.loc["Total Liabilities", balance]
    Main.loc[i, "Book Value"] = book_value
    try :
        P_B = (live_price * Sub_main_2.loc["Diluted Shares Outstanding", income]) / book_value
    except ZeroDivisionError :
        P_B = 0
    Main.loc[i, "P/B Ratio"] = P_B
    
    # Get Current Ratio
    try:
        current_ratio = Sub_main_2.loc["Total Current Assets", balance] / Sub_main_2.loc["Total Current Liabilities", balance]
    except ZeroDivisionError :
        current_ratio = 0
    Main.loc[i, "Current Ratio"] = current_ratio
    
    # Property, Plant & Equipment
    Main.loc[i, "PP&E"] = Sub_main_2.loc["Net Property, Plant & Equipment", balance]
    
    # Cash on Hand
    Main.loc[i, "Cash & Inv."] = Sub_main_2.loc["Cash & Short Term Investments", balance]
    
    # Revenue
    revenue = StockAnaPy.adding_data_past_data(Sub_main, Sub_main_2, "Sales/Revenue", income, delta_year, delta_quarter, y)
    Main.loc[i, "Revenue Yoy":"Rev 2020"] = revenue
    
    # Income
    profit = StockAnaPy.adding_data_past_data(Sub_main, Sub_main_2, "Net Income", income, delta_year, delta_quarter, y)
    Main.loc[i, "Income Yoy":"Income 2020"] = profit
    
    # Income / Revenue
    Main.loc[i, "Income/Rev"] = Main.loc[i, "Income Yoy"] / Main.loc[i, "Revenue Yoy"]
    
    # Revenue / Market Cap
    rev_marketcap = revenue[0] / Market_cap
    Main.loc[i, "Rev/MC"] = rev_marketcap
    
    
    # Book Value / Market Cap
    book_marketcap = book_value / Market_cap
    Main.loc[i, "Book/MC"] = book_marketcap
    
    # Capital Expenditure
    capex = StockAnaPy.adding_data_past_data(Sub_main, Sub_main_2, "Capital Expenditures", flow, delta_year, delta_quarter, y)
    Main.loc[i, "CAPEX":"CAPEX 2020"] = capex
    
    # Free Cash Flow
    fcf = StockAnaPy.adding_data_past_data(Sub_main, Sub_main_2, "Free Cash Flow", flow, delta_year, delta_quarter, y)
    Main.loc[i, "Free Cash Flow Yoy":"FCF 2020"] = fcf
    
    # Capital Expenditure / Revenue
    try :
        Main.loc[i, "CAPEX/Rev"] = -capex[0] / revenue[0]
    except ZeroDivisionError :
        Main.loc[i, "CAPEX/Rev"] = 0

    # PP&E / Assets ratio
    try :
        Main.loc[i, "PP&E/Assets"] = Main.loc[i, "PP&E"] / Sub_main_2.loc["Total Assets", balance]
    except ZeroDivisionError :
        Main.loc[i, "PP&E/Assets"] = 0
    
    # Cash / Assets ratio
    try :
        Main.loc[i, "Cash/Assets"] = Main.loc[i, "Cash & Inv."] / Sub_main_2.loc["Total Assets", balance]
    except ZeroDivisionError :
        Main.loc[i, "Cash/Assets"] = 0
    
    # Intangible Assets / Assets ratio
    try :
        Main.loc[i, "Int. Ass/Ass"] = Sub_main_2.loc["Intangible Assets", balance] / Sub_main_2.loc["Total Assets", balance]
    except ZeroDivisionError :
        Main.loc[i, "Int. Ass/Ass"] = 0
    
    # Debt / Shareholder Equity Ratio
    try :
        Main.loc[i, "Debt/Equity"] = (Sub_main_2.loc["Short Term Debt", balance] + Sub_main_2.loc["Long-Term Debt", balance]) / Sub_main_2.loc["Common Equity (Total)", balance]
    except ZeroDivisionError :
        Main.loc[i, "Debt/Equity"] = 0

    print (i, Main.loc[i, :])

    # Loop to the next ticker making sure we don't redownload data.
    i_2 = i_2 + 1

    # Creates a list of the metrics of the ticker being worked on. This will allow it to be copied directly onto Excel.
    Row = StockAnaPy.from_Main_to_Row(Main, i, combined_columns)

    # Clasifies the metrics into the right row of the right sheet.
    StockAnaPy.classify_into_sheet(Row, ind, file, i)

# Function to put all data into sheet 1 of excel in the other file.
StockAnaPy.put_into_sheet1(file, 'Companies', analysis_file)
exit()




