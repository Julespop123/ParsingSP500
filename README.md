# ParsingSP500
These scripts are a completely free way of downloading and comparing S&P 500 Companies.
The data used is parsed from MarketWatch. 
You must have Excel for this to work and must have Python as well as the necessary libraries.
To start out, download the Excel Spreadsheet templates. 
You can then run S&P500_File_Setup.py to initialize the Spreadhsheet. 
Then you run Parsing_Financial_Statements_SP500.py making sure the company_to_start_from = "MMM". You will need to change this to the latest ticker if a problem occurs while downloading the data.
Once all the data has downloaded, you can use Updating_stock_prices_SP.py script to update the file. This script is much faster as it does not redownload all the metrics again. Only downloads the new price. 
