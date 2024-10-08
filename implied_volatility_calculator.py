import numpy as np
from scipy.stats import norm
from scipy.optimize import minimize_scalar
import pandas as pd
import pandas_datareader.data as web
import datetime as dt

N = norm.cdf

# S : current asset price
# K: strike price of the option
# r: risk free rate
# T : time until option expiration
# σ: annualized volatility of the asset's returns


def BS_CALL(S, K, T, r, sigma):
    d1 = (np.log(S/K) + (r + sigma**2/2)*T) / (sigma*np.sqrt(T))
    d2 = d1 - sigma * np.sqrt(T)

    return S * N(d1) - K * np.exp(-r*T)* N(d2)


def BS_PUT(S, K, T, r, sigma):
    d1 = (np.log(S/K) + (r + sigma**2/2)*T) / (sigma*np.sqrt(T))
    d2 = d1 - sigma * np.sqrt(T)

    return K*np.exp(-r*T)*N(-d2) - S*N(-d1)


def implied_vol(opt_value, S, K, T, r, type_):
    try:
        def call_obj(sigma):
            return abs(BS_CALL(S, K, T, r, sigma) - opt_value)

        def put_obj(sigma):
            return abs(BS_PUT(S, K, T, r, sigma) - opt_value)

        if type_ == 'call':
            res = minimize_scalar(call_obj, bounds=(0.01, 6), method='bounded')
            return res.x
        elif type_ == 'put':
            res = minimize_scalar(put_obj, bounds=(0.01, 6), method='bounded')
            return res.x
        else:
            raise ValueError("type_ must be 'put' or 'call'")
    except Exception:
        raise


def main():
    nifty_19_sep_ce_price = 208
    nifty_19_sep_pe_price = 74

    spot_price = 25356
    strike_price = 25200
    time_to_expire = 4
    risk_free_rate = 0

    call_iv = implied_vol(
        nifty_19_sep_ce_price,
        spot_price,
        strike_price,
        time_to_expire,
        risk_free_rate,
        'call',
    )

    put_iv = implied_vol(
        nifty_19_sep_ce_price,
        spot_price,
        strike_price,
        time_to_expire,
        risk_free_rate,
        'put',
    )

    print(f'CE IV: {call_iv}')
    print(f'PE IV: {put_iv}')


if __name__ == '__main__':
    main()
