import os
from kiteconnect import KiteConnect
from typing import List
from openpyxl import load_workbook

API_KEY = 'p7qy2u03ev8e45pm'
API_SECRETE = '4pamm45xsirewovl8smza5t1qvft0t92'
ACCESS_TOKEN = None


def new_kite_connect_client() -> KiteConnect:
    kc: KiteConnect = KiteConnect(
        api_key=API_KEY,
    )

    print("Please login with here and fetch the 'request_token' from redirected "
          "url after successful login : ", kc.login_url())

    request_token: str = input("enter 'request_token': ")

    session_data: dict = kc.generate_session(
        request_token=request_token,
        api_secret=API_SECRETE,
    )

    ACCESS_TOKEN = session_data['access_token']
    kc.set_access_token(ACCESS_TOKEN)

    print('\nkite connect client creation successful !!! ')

    return kc


# output: {"10-20": [3, cum_occurrence, cum_%age], "20-30": [4, cum_occurrence, cum_%age]}
def get_distributions(
        nums: List[float],
        grouping_range_width: int,
) -> dict:
    nums.sort()
    momentum_distributions = {}

    start_range = 0
    end_range = grouping_range_width

    i = 0
    sum_of_occurrence = 0
    tot_cumulative_occurrence = 0

    while i < len(nums):
        if nums[i] <= end_range:
            sum_of_occurrence += 1
        else:
            key = f'{start_range}-{end_range}'
            momentum_distributions[key] = [sum_of_occurrence]
            tot_cumulative_occurrence += sum_of_occurrence

            start_range = end_range
            end_range = start_range + grouping_range_width

            sum_of_occurrence = 0
            i -= 1

        i += 1

    if sum_of_occurrence != 0:
        momentum_distributions[f'{start_range}-{end_range}'] = [sum_of_occurrence]
        tot_cumulative_occurrence += sum_of_occurrence

    # calculate cumulative percentage
    start_range = 0
    end_range = grouping_range_width
    tot_sum_so_far = 0
    while True:
        my_range = f'{start_range}-{end_range}'
        if my_range not in momentum_distributions:
            break

        count = momentum_distributions[my_range][0]
        tot_sum_so_far += count

        cur_percentage = round(tot_sum_so_far / tot_cumulative_occurrence * 100)
        momentum_distributions[my_range].append(tot_sum_so_far)
        momentum_distributions[my_range].append(cur_percentage)

        start_range = end_range
        end_range = start_range + grouping_range_width

    return momentum_distributions


def _clean_worksheet(ws, start_row: int, start_col: int):
    for row in range(start_row, 1000):  # Adjust range to include row 1000
        for col in range(start_col, start_col+5):  # Columns are typically 1-indexed
            cell = ws.cell(row=row, column=col)
            cell.value = None


def write_distribution_result_to_worksheet(
        distribution_data: dict,
        abs_file_path: str,
        sheet_name: str,
        start_row: int,
        start_col: int,
):
    wb = load_workbook(abs_file_path)
    ws = wb[sheet_name]

    _clean_worksheet(ws, start_row, start_col)

    cur_row = start_row
    for group_range, arr in distribution_data.items():
        count = arr[0]
        cum_sum = arr[1]
        cum_percentage = arr[2]

        ws.cell(row=cur_row, column=start_col, value=group_range)
        ws.cell(row=cur_row, column=start_col+1, value=count)
        ws.cell(row=cur_row, column=start_col+2, value=cum_sum)
        ws.cell(row=cur_row, column=start_col+3, value=cum_percentage)

        cur_row += 1

    wb.save(abs_file_path)
    wb.close()
