import pandas as pd
import numpy as np


def load_wb(wb_path: str):
    return pd.ExcelFile(wb_path)


def extract_company_data(df: pd.DataFrame):
    return {
        "ein": df.iloc[0, 0],
        "name": df.iloc[0, 1],
        "trade name": (lambda: "", lambda: df.iloc[0, 2])[type(df.iloc[0, 2]) is str](),
        "address": df.iloc[0, 3],
        "phone": df.iloc[0, 4],
        "city": df.iloc[0, 5],
        "state": df.iloc[0, 6],
        "zip": (lambda: int(df.iloc[0, 7]), lambda: df.iloc[0, 7])[
            type(df.iloc[0, 7]) is str
        ](),  # calling int here gets rid of the decimal .0
    }


def extract_tax_data(df: pd.DataFrame, row: int, round_delta: float) -> dict:
    return {
        "18a": excel_round(df.iloc[row, 3], round_delta),
        "26a": excel_round(df.iloc[row + 2, 3], round_delta),
        "27": excel_round(df.iloc[row + 4, 3], round_delta),
        "30": excel_round(df.iloc[row + 6, 3], round_delta),
    }


# This seems more accurate more of the time
def excel_round(num: np.float64, round_delta: float):
    num = np.round(num, 3)
    if np.floor(num * 1000) % 5 == 0:
        return round(num + round_delta, 2)
    return round(num, 2)


def add_commas_to_dollars(num: int):
    step = 3
    out = []
    current = str(num)
    while len(current) > step:
        out.append(current[-1 * step :])
        current = current[: len(current) - step]
    out.append(current)
    out.reverse()
    return ",".join(out)


def extract_dollars_and_cents(num: np.float63) -> list:
    if int(np.round(num)) == -1:
        return ["", ""]
    dollars = int(np.floor(num))
    # dollars = add_commas_to_dollars(dollars)
    cents = str(num)[-3:]  # Don't forget the colon
    if cents[-1] == ".":
        cents = f"{cents[0]}0"
    return [dollars, cents]


def load_data(wb: pd.ExcelFile, sheets: dict, row_2020: int, row_2021: int, round_delta: float):
    return {
        "company": extract_company_data(wb.parse(sheet_name=sheets["input"])),
        "2020_q2": extract_tax_data(
            wb.parse(sheet_name=sheets["2020Q2"]), row_2020, round_delta
        ),
        "2020_q3": extract_tax_data(
            wb.parse(sheet_name=sheets["2020Q3"]), row_2020, round_delta
        ),
        "2020_q4": extract_tax_data(
            wb.parse(sheet_name=sheets["2020Q4"]), row_2020, round_delta
        ),
        "2021_q1": extract_tax_data(
            wb.parse(sheet_name=sheets["2021Q1"]), row_2021, round_delta
        ),
        "2021_q2": extract_tax_data(
            wb.parse(sheet_name=sheets["2021Q2"]), row_2021, round_delta
        ),
        "2021_q3": extract_tax_data(
            wb.parse(sheet_name=sheets["2021Q3"]), row_2021, round_delta
        ),
    }
