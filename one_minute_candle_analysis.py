import json
from datetime import date, time, datetime, timedelta
from typing import List
import requests

end_datetime = datetime.today()
start_datetime = end_datetime - timedelta(days=180)

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


class UpstoxCandlesticksData:
    def __init__(self, candles: List[Candle]):
        self.candles: List[Candle] = candles


class UpstoxCandlestickResponse:
    def __init__(self, upstox_candlesticks_data: UpstoxCandlesticksData):
        self.status = 'success'
        self.data = upstox_candlesticks_data

    @classmethod
    def from_upstox_api_response(cls, resp: dict):
        print(f'resp: {resp}')
        return UpstoxCandlestickResponse(
            upstox_candlesticks_data=UpstoxCandlesticksData(
                [Candle.from_api_resp_candle_dict(candle_dict) for candle_dict in resp['data']['candles']]
            ),
        )


def fetch_candlestick_data_from_upstox() -> UpstoxCandlestickResponse:
    resp = requests.get(
        url='https://api.upstox.com/v2/historical-candle/NSE_INDEX%7CNifty%2050/1minute/{}/{}'
        .format(end_date.strftime(date_str_format), start_date.strftime(date_str_format))
    )

    if resp.status_code != 200:
        raise Exception(f'Upstox fetch api failed, status_code: {resp.status_code}')

    return UpstoxCandlestickResponse.from_upstox_api_response(json.loads(resp.content))


def main():
    ...


if __name__ == '__main__':
    main()
