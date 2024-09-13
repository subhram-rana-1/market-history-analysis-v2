## How to fetch historical NIFTY and BANKNIFTY data
* Change `market_instrument_token` in `config.py` file
* Update `report_file_relative_path = 'BANKNIFTY_historical_analysis.xlsx'` in `fetch_historical_data.py` file
* Run `python fetch_historical_data.py`
  * login in kite website as per the instruction and link given in terminal
* Get the report in `BANKNIFTY_historical_analysis.xlsx` file's `daily_fluctuation` sheet

## How to generate daily fluctuation occurrence data for NIFTY and BANKNIFTY
* Set out pur report file name in `occurrence_distribution.py` file, Ex: report_file_relative_path = 'NIFTY_historical_analysis.xlsx'
* Update `end_row` as per `daily_fluctuation` sheet in `BANKNIFTY_historical_analysis.xlsx` file 
* Run `python occurrence_distribution.py`
* Get report in `BANKNIFTY_historical_analysis.xlsx` file's `fluctuation_distribution` sheet.
