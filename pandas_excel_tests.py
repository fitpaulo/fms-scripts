import pandas as pd

PAYROLL_COLS = ["2020 Q2", "2020 Q3", "2020 Q4", "2021 Q1", "2021 Q2", "2021 Q3"]


def new_pay_dict():
    return {
        "2020 Q2": "",
        "2020 Q2 Health": "",
        "2020 Q3": "",
        "2020 Q3 Health": "",
        "2020 Q4": "",
        "2020 Q4 Health": "",
        "2021 Q1": "",
        "2021 Q1 Health": "",
        "2021 Q2": "",
        "2021 Q2 Health": "",
        "2021 Q3": "",
        "2021 Q3 Health": "",
    }


if __name__ == "__main__":
    file = "C:\\Users\\dguim\\OneDrive\\Documents\\src\\fml\\IES KENTUCKY ERTC Worksheet.xlsx"
    df = pd.ExcelFile(file)
    pay_ws = df.parse(sheet_name="Payroll worksheet")
    df_pay_dict = pay_ws.to_dict()
    rows = pay_ws.shape[0]  # returns tuple (rows, cols)
    payment_data = {}

    ignored_names = ["total", "diff"]

    print(pay_ws.head())
    print(pay_ws.iloc[1, 2])
    # print(pay_ws["2020 Q2 HEALTH"].isnull())

    # skip = False
    # for i in range(0, rows):
    #     print(df_pay_dict["EMPLOYEE"][i])
    #     if type(df_pay_dict["EMPLOYEE"][i]) == float:  # nan is a "float"
    #         continue
    #     for ignored_name in ignored_names:
    #         if ignored_name in df_pay_dict["EMPLOYEE"][i]:
    #             skip = True
    #     if skip:
    #         skip = False
    #         continue
