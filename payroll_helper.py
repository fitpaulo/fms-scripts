from pprint import pp
import pandas as pd
import yaml


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
        payroll_helper.load_data()
        # pp(payroll_helper.output_data)
        pp(list(payroll_helper.output_data))
        payroll_helper.create_output_df()
