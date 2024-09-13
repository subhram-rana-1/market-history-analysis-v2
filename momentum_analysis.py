import os
from typing import List
from openpyxl import load_workbook


input_sheet_relative_file_path = 'NIFTY_historical_analysis.xlsx'
momentum_sheet_name = 'momentum_analysis'
momentum_col = 2
start_row = 3
end_row = 121
grouping_range_width = 10


def fetch_momentums_from_input_sheet() -> List[int]:
    full_path = os.getcwd() + '/' + input_sheet_relative_file_path
    wb = load_workbook(full_path)
    ws = wb[momentum_sheet_name]

    momentums = []
    for row in ws.iter_rows(min_col=momentum_col, max_col=momentum_col, min_row=start_row, max_row=end_row):
        for cell in row:
            momentums.append(cell.value)

    return momentums


# output: {"10-20": [3, cum_occurrence, cum_%age], "20-30": [4, cum_occurrence, cum_%age]}
def get_momentum_distributions(momentums: List[int]) -> dict:
    momentums.sort()
    momentum_distributions = {}

    start_range = 0
    end_range = grouping_range_width

    i = 0
    sum_of_occurrence = 0
    tot_cumulative_occurrence = 0

    while i < len(momentums):
        if momentums[i] <= end_range:
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


def _clean_worksheet(ws):
    for row in range(3, 1000):  # Adjust range to include row 1000
        for col in [5, 6, 7, 8]:  # Columns are typically 1-indexed
            cell = ws.cell(row=row, column=col)
            cell.value = None


def write_output_to_sheet(momentum_distributions: dict):
    full_path = os.getcwd() + '/' + input_sheet_relative_file_path

    wb = load_workbook(full_path)
    ws = wb[momentum_sheet_name]

    _clean_worksheet(ws)

    cur_row = 3
    for group_range, arr in momentum_distributions.items():
        count = arr[0]
        cum_sum = arr[1]
        cum_percentage = arr[2]

        ws.cell(row=cur_row, column=5, value=group_range)
        ws.cell(row=cur_row, column=6, value=count)
        ws.cell(row=cur_row, column=7, value=cum_sum)
        ws.cell(row=cur_row, column=8, value=cum_percentage)

        cur_row += 1

    wb.save(full_path)
    wb.close()


def main():
    momentums = fetch_momentums_from_input_sheet()
    momentum_distributions = get_momentum_distributions(momentums)
    write_output_to_sheet(momentum_distributions)


if __name__ == '__main__':
    main()
