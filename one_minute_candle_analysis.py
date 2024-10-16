import os
from datetime import date, datetime, timedelta
from common import write_distribution_result_to_worksheet, fetch_candlestick_data_from_upstox

# User input -----------------------------
NIFTY_INSTRUMENT_TOKEN = "NSE_INDEX%7CNifty%2050"

end_datetime = datetime.today()
start_datetime = end_datetime - timedelta(days=180)
unique_instrument_token = NIFTY_INSTRUMENT_TOKEN
sheet_rel_file_path = 'one_minute_candle_analysis_report.xlsx'
up_sheet_name = 'open-to-high'
down_sheet_name = 'open-to-low'
body_sheet_name = 'body'
lower_wick_sheet_name = 'lower-wick'
upper_wick_sheet_name = 'upper-wick'
bucket_size = 2
max_candle_size_for_analysis = 40
# User input -----------------------------


# -------------- CODE ------------------------------------------------
# --------------------------------------------------------------------

start_date: date = start_datetime.date()
end_date: date = end_datetime.date()


def write_to_sheet(move_distributions: dict, sheet_name: str):
    full_path = os.getcwd() + '/' + sheet_rel_file_path
    write_distribution_result_to_worksheet(
        move_distributions,
        full_path,
        sheet_name,
        2,
        1,
    )


def main():
    upstox_1min_candlestick_resp = fetch_candlestick_data_from_upstox(
        unique_instrument_token,
        start_date,
        end_date,
    )

    up_move_distribution = \
        upstox_1min_candlestick_resp.distribution_for_up_moves(bucket_size, max_candle_size_for_analysis)
    down_move_distribution = \
        upstox_1min_candlestick_resp.distribution_for_down_moves(bucket_size, max_candle_size_for_analysis)
    candle_body_distribution = \
        upstox_1min_candlestick_resp.distribution_for_candle_body(bucket_size, max_candle_size_for_analysis)
    candle_lower_wick_distribution = \
        upstox_1min_candlestick_resp.distribution_for_candle_lower_wick(bucket_size, max_candle_size_for_analysis)
    candle_upper_wick_distribution = \
        upstox_1min_candlestick_resp.distribution_for_candle_upper_wick(bucket_size, max_candle_size_for_analysis)

    write_to_sheet(up_move_distribution, up_sheet_name)
    write_to_sheet(down_move_distribution, down_sheet_name)
    write_to_sheet(candle_body_distribution, body_sheet_name)
    write_to_sheet(candle_lower_wick_distribution, lower_wick_sheet_name)
    write_to_sheet(candle_upper_wick_distribution, upper_wick_sheet_name)


if __name__ == '__main__':
    main()
