import pandas as pd
import numpy as np


class excel_helper:
    """This class is specifically intended to help with working with the WS used by FMS."""

    def __init__(
        self,
        wb_path: str,
        sheet_names: dict,
        round_delta: float,
        row_2020: int,
        row_2021: int,
        col: int = 3,
    ):
        """This is the constructor for excel_helper.

        Args:
            wb_path (str): The full path to the Excel File /path/to/file.xlsx
            sheet_names (dict): A dict of the names of the sheets in the WB
            round_delta (float): This is used when rounding excel values
            row_2020 (int): The integer row number of the first tax item 18a
            row_2021 (int): The integer row number of 18a for 2021 sheets
            col (int): The col for 18a, default is 3 (D in excel)
        """
        self.wb = pd.ExcelFile(wb_path)
        self.sheet_names = sheet_names
        self.round_delta = round_delta
        self.row_2020 = row_2020
        self.row_2021 = row_2021
        self.col = col

    def extract_company_data(self, df: pd.DataFrame) -> dict:
        """This function extracts the company data from the Data Input page

        Args:
            df (pd.DataFrame): This needs to be the Data Input page

        Returns:
            dict: A dict of the company data
        """
        return {
            "ein": df.iloc[0, 0],
            "name": df.iloc[0, 1],
            "trade name": (lambda: "", lambda: df.iloc[0, 2])[
                type(df.iloc[0, 2]) is str
            ](),
            "address": df.iloc[0, 3],
            "phone": df.iloc[0, 4],
            "city": df.iloc[0, 5],
            "state": df.iloc[0, 6],
            "zip": (lambda: int(df.iloc[0, 7]), lambda: df.iloc[0, 7])[
                type(df.iloc[0, 7]) is str
            ](),  # calling int here gets rid of the decimal .0
        }

    def extract_tax_data(self, df: pd.DataFrame, row: int) -> dict:
        """This specically extracts the tax data from a tax sheet

        Args:
            df (pd.DataFrame): The specific tax sheet
            row (int): The integer of the row

        Returns:
            dict: The data from the sheet in dict format
        """
        return {
            "18a": self.excel_round(df.iloc[row, self.col], self.round_delta),
            "26a": self.excel_round(df.iloc[row + 2, self.col], self.round_delta),
            "27": self.excel_round(df.iloc[row + 4, self.col], self.round_delta),
            "30": self.excel_round(df.iloc[row + 6, self.col], self.round_delta),
        }

    def excel_round(self, num: np.float64) -> float:
        """Oh boy!  What to say here.  Rounding is a very difficult thing with floats.
        Numpy isn't good enough.  We can do better by adding a small delta when 3rd decimal
        is a 5 (or a 0, but that won't affect anything)

        Args:
            num (np.float64): The number we are rounding

        Returns:
            _type_: _description_
        """
        num = np.round(num, 3)
        if np.floor(num * 1000) % 5 == 0:
            return round(num + self.round_delta, 2)
        return round(num, 2)

    def extract_dollars_and_cents(num: np.float63) -> list:
        """Separate the decimal and the whole number

        Args:
            num (np.float64): The number from the excel workbook as a float

        Returns:
            list: A list of the two pats [whole, decimal] [dollars, cents]
        """
        if int(np.round(num)) == -1:
            return ["", ""]
        dollars = int(np.floor(num))
        # dollars = add_commas_to_dollars(dollars)
        cents = str(num)[-3:]  # Don't forget the colon
        if cents[-1] == ".":
            cents = f"{cents[0]}0"
        return [dollars, cents]

    def load_data(self) -> dict:
        """Run through each sheet and parse the data from it

        Args:
            wb (pd.ExcelFile): The excel workbook
            sheets (dict): A dict of the page names
            row_2020 (int): The row to get tax data in 2020
            row_2021 (int): The row to get tax data in 2021
            round_delta (float): A delta to add when rounding

        Returns:
            dict: A dict with all the data we extracted
        """
        self.data = {
            "company": self.extract_company_data(
                self.wb.parse(sheet_name=self.sheets["input"])
            ),
            "2020_q2": self.extract_tax_data(
                self.wb.parse(sheet_name=self.sheets["2020Q2"]),
                self.row_2020,
                self.round_delta,
            ),
            "2020_q3": self.extract_tax_data(
                self.wb.parse(sheet_name=self.sheets["2020Q3"]),
                self.row_2020,
                self.round_delta,
            ),
            "2020_q4": self.extract_tax_data(
                self.wb.parse(sheet_name=self.sheets["2020Q4"]),
                self.row_2020,
                self.round_delta,
            ),
            "2021_q1": self.extract_tax_data(
                self.wb.parse(sheet_name=self.sheets["2021Q1"]),
                self.row_2021,
                self.round_delta,
            ),
            "2021_q2": self.extract_tax_data(
                self.wb.parse(sheet_name=self.sheets["2021Q2"]),
                self.row_2021,
                self.round_delta,
            ),
            "2021_q3": self.extract_tax_data(
                self.wb.parse(sheet_name=self.sheets["2021Q3"]),
                self.row_2021,
                self.round_delta,
            ),
        }

    def fix_zip(self):
        """Fix zip code if it was an integer starting with 0, we want that 0 to be there
        """
        if self.data["company"]["zip"] is int:
            if self.data["company"]["zip"] < 10000:
                self.data["company"]["zip"] = f"0{self.data['company']['zip']}"
