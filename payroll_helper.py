from pprint import pp
import os
import sys
import pandas as pd
import numpy as np
import yaml
import re


class PayrollHelper:
    def __init__(
        self,
        path: str,
        filename: str,
        output_filename: str,
        base_data_header: str,
        index_col: int,
        header: int,
    ) -> None:
        self.path = path
        self.filename = filename
        self.output_filename = output_filename
        self.base_data_header = base_data_header
        self.index_col = index_col
        self.header = header
        self.output_data = {}
        self.load_wb()
        self.parse_output_index()

    def load_wb(self):
        self.wb = pd.ExcelFile(f"{self.path}\\{self.filename}")
        self.df = self.wb.parse()

    def parse_output_index(self):
        df = self.parse_first_df()
        self.output_index = df.index.tolist()

    def parse_df(
        self, sheet_name: str, index_col: int = None, header: int = None
    ) -> pd.DataFrame:
        return self.wb.parse(
            sheet_name=sheet_name,
            index_col=index_col or self.index_col,
            header=header or self.header,
        )

    def parse_first_df(self) -> pd.DataFrame:
        return self.parse_df(self.wb.sheet_names[-1])

    def name_is_valid(self, name):
        name = name.strip()  # make sure it is clean
        ellipsis = "..."
        ellipsis2 = "…"
        if ellipsis in name or ellipsis2 in name:
            return False
        if "" == name:
            return False
        if "unnamed" in name.lower():
            return False
        return True

    def process_multi_col(self, df: pd.DataFrame) -> bool:
        processed = False
        ignore_list = ["unnamed"]
        for ignored in ignore_list:
            if ignored in df.columns[0].lower():
                for col in df.columns:
                    if ignored not in col.lower():
                        if self.name_is_valid(col):
                            self.output_data[col.strip()] = None
                            processed = True
        return processed

    def process_single_col(self, df: pd.DataFrame) -> bool:
        processed = False
        for emp in df.columns[0].split("  "):
            if len(emp) > -1:
                if self.name_is_valid(emp):
                    self.output_data[emp.strip()] = None
                    processed = True
        return processed

    def load_employee_names(self):
        for sheet in self.wb.sheet_names:
            df = self.wb.parse(sheet_name=sheet)
            if not self.process_multi_col(df):
                self.process_single_col(df)
        self.output_data["TOTAL"] = None

    def load_data_column(
        self,
        df: pd.DataFrame,
        idx: int,
        col_header_key: str,
    ) -> bool:
        success = True
        try:
            self.output_data[self.employees[idx]] = df[col_header_key].tolist()
        except KeyError:  # This header is not in the current df
            success = False
        except IndexError:
            raise "Names are missing!"
        return success

    def load_employee_data(self):
        self.employees = list(self.output_data.keys())
        idx = 0
        col_header_keys = [
            self.base_data_header,
            f"{self.base_data_header}.1",
            f"{self.base_data_header}.2",
        ]
        for sheet in self.wb.sheet_names:
            df = self.parse_df(sheet_name=sheet)
            for col_header_key in col_header_keys:
                if self.load_data_column(df, idx, col_header_key):
                    idx += 1

    def load_data(self):
        self.load_employee_names()
        self.load_employee_data()

    def create_output_df(self):
        output_df = pd.DataFrame(index=self.output_index, columns=self.employees)
        for name, salary_data in self.output_data.items():
            self.normalize_data(salary_data)
            output_df[name] = salary_data
        output_df.to_excel(self.output_filename)

    def normalize_data(self, data: list):
        diff = len(self.output_index) - len(data)
        for i in range(0, diff):
            data.append(0)


def extract_gross(row):
    gross_col = 5
    print(row[gross_col])
    return float(row[gross_col])


def extract_gross_with_title(row):
    title_col = 0
    gross_col = 4
    try:
        if "totals" in row[title_col].lower():
            return float(row[gross_col])
    except AttributeError:
        pass
    return 0


def extract_ss(row):
    text_col = 7
    num_col = 8
    try:
        if row[text_col].lower() == "fica-ss":
            return float(row[num_col])
        elif row[text_col].lower() == "fed socsec - reclac":
            return float(row[num_col])
        return 0
    except AttributeError:
        return 0


if __name__ == "__main__":
    with open("conf/payroll.yaml") as file:
        conf = yaml.safe_load(file)
        payroll_helper = PayrollHelper(
            path=conf["path"],
            filename=conf["filename"],
            output_filename=conf["outfile"],
            base_data_header=conf["base_data_header"],
            index_col=conf["index_col"],
            header=conf["header"],
        )
        names = []
        gross = []
        ss = []
        payroll_helper.load_data()
        # pp(payroll_helper.df.to_string())
        cols = payroll_helper.df.columns.tolist()
        find_ss = False
        find_ss_recalc = False
        find_gross = False
        current_gross = 0
        current_ss = 0
        name_row = 3
        for _, i in payroll_helper.df.iterrows():
            if find_ss:
                result = extract_ss(i)
                if result > 0:
                    ss.append(result)
                    find_ss = False
            #     if find_ss_recalc:
            #         if result > 0:
            #             current_ss += result
            #         else:
            #             find_ss = False
            #             find_ss_recalc = False
            #             ss.append(current_ss)
            #     else:
            #         if result > 0:
            #             current_ss = result
            #             find_ss_recalc = True
            if find_gross:
                result = extract_gross_with_title(i)
                if result > 0:
                    gross.append(result)
                    find_gross = False
                # result = extract_gross(i)
                # if result is np.nan:
                    # find_gross = False
                    # gross.append(current_gross)
                # else:
                    # current_gross = result
            if type(i[name_row]) is str:
                # if re.search(r"^\w+, \w+", i[name_row]):
                if "," in i[name_row]:
                    if find_ss and find_gross:
                        # employee had 0 income, remove them
                        names.pop()
                    # if find_ss:  # 1099 emp didn't have
                    #     ss.append(0)
                    names.append(i[name_row])
                    find_ss = True
                    find_gross = True
        data = {"Name": names, "Gross": gross, "SocSec": ss}
        # pp(data)
        # print(len(data["Gross"]), len(data["Name"]), len(data["SocSec"]))
        df = pd.DataFrame(data)
        # pp(df.to_string())
        df.to_excel(os.path.join(conf["path"], conf["outfile"]))
