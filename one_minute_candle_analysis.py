import json
import os
from datetime import date, time, datetime, timedelta
from typing import List
import requests

from common import get_distributions, write_distribution_result_to_worksheet

# User input -----------------------------
NIFTY_INSTRUMENT_TOKEN = "NSE_INDEX%7CNifty%2050"

end_datetime = datetime.today()
start_datetime = end_datetime - timedelta(days=180)
unique_instrument_token = NIFTY_INSTRUMENT_TOKEN
sheet_rel_file_path = 'one_minute_candle_analysis_report.xlsx'
up_sheet_name = 'open-to-high'
down_sheet_name = 'open-to-low'
bucket_size = 2
max_candle_size_for_analysis = 40
# User input -----------------------------


# -------------- CODE ------------------------------------------------
# --------------------------------------------------------------------

start_date: date = start_datetime.date()
end_date: date = end_datetime.date()

upstox_ts_format = "%Y-%m-%dT%H:%M:%S%z"
date_str_format = "%Y-%m-%d"
time_str_format = "%H:%M:%S"
datetime_str_format = "%Y-%m-%d %H:%M:%S"


class Candle:
    def __init__(
            self,
            ts: datetime,
            open: float,
            hi: float,
            lo: float,
            close: float,
    ):
        self.ts = ts
        self.open = open
        self.hi = hi
        self.lo = lo
        self.close = close

    @classmethod
    def from_api_resp_candle_dict(cls, candle_dict: dict):
        return Candle(
            datetime.strptime(candle_dict[0], upstox_ts_format),
            candle_dict[1],
            candle_dict[2],
            candle_dict[3],
            candle_dict[4],
        )

    @property
    def up_move(self) -> float:
        return self.hi - self.open

    @property
    def down_move(self) -> float:
        return self.open - self.lo


class UpstoxCandlesticksData:
    def __init__(self, candles: List[Candle]):
        self.candles: List[Candle] = candles

    def get_moves(self, direction: str) -> List[float]:
        if direction == 'up':
            return [candle.up_move for candle in self.candles]
        if direction == 'down':
            return [candle.down_move for candle in self.candles]


class UpstoxCandlestickResponse:
    def __init__(self, upstox_candlesticks_data: UpstoxCandlesticksData):
        self.status = 'success'
        self.data = upstox_candlesticks_data

    @classmethod
    def from_upstox_api_response(cls, resp: dict):
        return UpstoxCandlestickResponse(
            upstox_candlesticks_data=UpstoxCandlesticksData(
                [Candle.from_api_resp_candle_dict(candle_dict) for candle_dict in resp['data']['candles']]
            ),
        )

    def distribution_for_moves(self, direction: str, bucket_length: int) -> dict:
        moves = self.data.get_moves(direction)

        # IMPORTANT - remove too big candles, they are outliers to this analysis
        moves = [x for x in moves if x <= max_candle_size_for_analysis]

        return get_distributions(moves, bucket_length)

    def distribution_for_up_moves(self, bucket_length: int) -> dict:
        return self.distribution_for_moves('up', bucket_length)

    def distribution_for_down_moves(self, bucket_length: int) -> dict:
        return self.distribution_for_moves('down', bucket_length)


def fetch_candlestick_data_from_upstox() -> UpstoxCandlestickResponse:
    resp = requests.get(
        url='https://api.upstox.com/v2/historical-candle/{}/1minute/{}/{}'
        .format(
            unique_instrument_token,
            end_date.strftime(date_str_format),
            start_date.strftime(date_str_format),
        )
    )

    if resp.status_code != 200:
        raise Exception(f'Upstox fetch api failed, status_code: {resp.status_code}')

    return UpstoxCandlestickResponse.from_upstox_api_response(json.loads(resp.content))


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
    upstox_1min_candlestick_resp = fetch_candlestick_data_from_upstox()

    up_move_distribution = upstox_1min_candlestick_resp.distribution_for_up_moves(bucket_size)
    down_move_distribution = upstox_1min_candlestick_resp.distribution_for_down_moves(bucket_size)

    write_to_sheet(
        up_move_distribution,
        up_sheet_name,
    )

    write_to_sheet(
        down_move_distribution,
        down_sheet_name,
    )


if __name__ == '__main__':
    main()
