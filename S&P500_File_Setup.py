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
StockAnaPy.classify_by_sheet(tickers, 1, file, combined_columns)
exit()
