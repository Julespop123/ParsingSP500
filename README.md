# ParsingSP500
These scripts are a completely free way of downloading and comparing S&P 500 Companies.
The data used is parsed from MarketWatch. 
You must have Excel for this to work and must have Python as well as the necessary libraries.
Follow these steps to use it on your local repository:

 - Download the Excel Spreadsheet template and all the Python Scripts, store them in your local repository. In addition, create a blank Excel Spreadhshet titled SP_500_Companies.xlsx.
 - You can then run S&P500_File_Setup.py to initialize the Spreadhsheet. This will load all the current S&P500 companies into the SP_500_Companies spreadsheet.
 - Then you run Parsing_Financial_Statements_SP500.py making sure the company_to_start_from = "MMM". You will need to change this to the latest ticker downloaded if a problem occurs while downloading the data.
 - Once all the data has downloaded, you can use Updating_stock_prices_SP.py script to update the file. This script is much faster as it does not redownload all the metrics again. Only downloads the new price. 

Enjoy!

Libraries that are used are : OpenPyxl, YFinance, BeautifulSoup, datetime, Pandas, Numpy, Yahoo_Fin.
