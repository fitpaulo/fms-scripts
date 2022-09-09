import PyPDF2
from PyPDF2.generic import NameObject
import pandas as pd
import numpy as np
import yaml
import datetime
import os


with open("conf/config.yaml", "r") as file:
    conf = yaml.safe_load(file)
with open("conf/f941x.yaml", "r") as file:
    pdf_conf = yaml.safe_load(file)

# From yaml
COMPANY_PATH = conf["path"]
COPANY_TYPE = conf["type"]
WS_NAME = conf["ws"]
SKIP_8821 = conf["skip"]
YEAR_QUARTER = conf["year_quarter"]
DROPBOX_PATH = conf["dropbox_path"]
PDF_DICT = pdf_conf["pdf_dict"]
PAYROLL_DIR = conf["payroll_dir"]
QUARTER_FIELDS = pdf_conf["quarter_fields"]
SHEETS = conf["excel_sheet_names"]
ROUNND_DELTA = conf["round_delta"]
ROW_2020 = conf["row_2020"]
ROW_2021 = conf["row_2021"]

# Dynamic vars
BASE_PATH = f"{DROPBOX_PATH}\\COMPANIES {COPANY_TYPE}"
PDF_PATH = f"{BASE_PATH}\\PAT {COPANY_TYPE} ERTC"
F941X_PATH = f"{PDF_PATH}\\f941x 8-9-22.pdf"
F8821_PATH = f"{PDF_PATH}\\f8821 8-9-22.pdf"
WS_PATH = f"{BASE_PATH}\\{COMPANY_PATH}\\{PAYROLL_DIR}\\{WS_NAME}.xlsx"
OUTPUT_PATH = f"{BASE_PATH}\\{COMPANY_PATH}\\941x"
NEWLINE = os.linesep


def print_non_empty_fields(pdf: PyPDF2.PdfFileReader):
    field_dict = pdf.getFormTextFields()
    for k, v in field_dict.items():
        if v:
            print(f"key:{k}   value:{v}")


def update_quater_check_box(page: PyPDF2._page.PageObject, value: int):
    for i in range(0, len(page["/Annots"])):
        annot = page["/Annots"][i].getObject()
        for k, field in QUARTER_FIELDS.items():
            if annot.get("/T") == field:
                as_value = "/Off"
                if k == value:
                    as_value = f"/{value}"
                annot.update(
                    {
                        NameObject("/V"): NameObject(as_value),
                        NameObject("/AS"): NameObject(as_value),
                    }
                )


# This is needed to update checkboxes
def get_page_object_data(page: PyPDF2._page.PageObject):
    for i in range(0, len(page["/Annots"])):
        tmp = page["/Annots"][i].get_object()
        if "c1_" in tmp.get("/T"):
            # from pprint import pprint as pp
            # pp(tmp)
            pass


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


def extract_tax_data(df: pd.DataFrame, row: int) -> dict:
    return {
        "18a": excel_round(df.iloc[row, 3]),
        "26a": excel_round(df.iloc[row + 2, 3]),
        "27": excel_round(df.iloc[row + 4, 3]),
        "30": excel_round(df.iloc[row + 6, 3]),
    }


# This seems more accurate more of the time
def excel_round(num):
    num = np.round(num, 3)
    if np.floor(num * 1000) % 5 == 0:
        return round(num + ROUNND_DELTA, 2)
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


def extract_dollars_and_cents(num: np.float64) -> list:
    if int(np.round(num)) == 0:
        return ["", ""]
    dollars = int(np.floor(num))
    # dollars = add_commas_to_dollars(dollars)
    cents = str(num)[-2:]  # Don't forget the colon
    if cents[0] == ".":
        cents = f"{cents[1]}0"
    return [dollars, cents]


def write_pdf_data(pdf: PyPDF2.PdfFileWriter, data: dict, year: int, quarter: int):
    d = datetime.date.today()
    dollars_18a, cents_18a = extract_dollars_and_cents(
        data[f"{year}_q{quarter}"]["18a"]
    )
    dollars_26a, cents_26a = extract_dollars_and_cents(
        data[f"{year}_q{quarter}"]["26a"]
    )
    dollars_27, cents_27 = extract_dollars_and_cents(data[f"{year}_q{quarter}"]["27"])
    dollars_30, cents_30 = extract_dollars_and_cents(data[f"{year}_q{quarter}"]["30"])
    pdf_writer.update_page_form_field_values(
        pdf_writer.pages[0],
        {
            PDF_DICT["p1"]["ein1"]: data["company"]["ein"][:2],
            PDF_DICT["p1"]["ein2"]: data["company"]["ein"][3:],
            PDF_DICT["p1"]["name"]: data["company"]["name"],
            PDF_DICT["p1"]["trade name"]: data["company"]["trade name"],
            PDF_DICT["p1"]["address"]: data["company"]["address"],
            PDF_DICT["p1"]["city"]: data["company"]["city"],
            PDF_DICT["p1"]["state"]: data["company"]["state"],
            PDF_DICT["p1"]["zip"]: data["company"]["zip"],
            PDF_DICT["p1"]["city"]: data["company"]["city"],
            PDF_DICT["p1"]["yyyy"]: d.year,
            PDF_DICT["p1"]["dd"]: f"{d:%d}",
            PDF_DICT["p1"]["mm"]: f"{d:%m}",
            PDF_DICT["p1"]["year correcting"]: year,
        },
    )
    pdf_writer.update_page_form_field_values(
        pdf_writer.pages[1],
        {
            PDF_DICT["p2"]["ein"]: data["company"]["ein"],
            PDF_DICT["p2"]["name"]: data["company"]["name"],
            PDF_DICT["p2"]["year"]: year,
            PDF_DICT["p2"]["quarter"]: quarter,
            PDF_DICT["tax"]["18a_dollars"][0]: dollars_18a,
            PDF_DICT["tax"]["18a_dollars"][1]: dollars_18a,
            PDF_DICT["tax"]["18a_dollars"][2]: -1 * dollars_18a,
            PDF_DICT["tax"]["18a_cents"][0]: cents_18a,
            PDF_DICT["tax"]["18a_cents"][1]: cents_18a,
            PDF_DICT["tax"]["18a_cents"][2]: cents_18a,
        },
    )
    pdf_writer.update_page_form_field_values(
        pdf_writer.pages[2],
        {
            PDF_DICT["p3"]["ein"]: data["company"]["ein"],
            PDF_DICT["p3"]["name"]: data["company"]["name"],
            PDF_DICT["p3"]["year"]: year,
            PDF_DICT["p3"]["quarter"]: quarter,
            PDF_DICT["tax"]["23_dollars"]: -1 * dollars_18a,
            PDF_DICT["tax"]["23_cents"]: cents_18a,
            PDF_DICT["tax"]["26a_dollars"][0]: dollars_26a,
            PDF_DICT["tax"]["26a_dollars"][1]: dollars_26a,
            PDF_DICT["tax"]["26a_dollars"][2]: -1 * dollars_26a,
            PDF_DICT["tax"]["26a_cents"][0]: cents_26a,
            PDF_DICT["tax"]["26a_cents"][1]: cents_26a,
            PDF_DICT["tax"]["26a_cents"][2]: cents_26a,
            PDF_DICT["tax"]["27_dollars"]: -1 * dollars_27,
            PDF_DICT["tax"]["27_cents"]: cents_27,
            PDF_DICT["tax"]["30_dollars"][0]: dollars_30,
            PDF_DICT["tax"]["30_dollars"][1]: dollars_30,
            PDF_DICT["tax"]["30_cents"][0]: cents_30,
            PDF_DICT["tax"]["30_cents"][1]: cents_30,
        },
    )
    pdf_writer.update_page_form_field_values(
        pdf_writer.pages[3],
        {
            PDF_DICT["p4"]["ein"]: data["company"]["ein"],
            PDF_DICT["p4"]["name"]: data["company"]["name"],
            PDF_DICT["p4"]["year"]: year,
            PDF_DICT["p4"]["quarter"]: quarter,
        },
    )
    pdf_writer.update_page_form_field_values(
        pdf_writer.pages[4],
        {
            PDF_DICT["p5"]["ein"]: data["company"]["ein"],
            PDF_DICT["p5"]["name"]: data["company"]["name"],
            PDF_DICT["p5"]["year"]: year,
            PDF_DICT["p5"]["quarter"]: quarter,
        },
    )


def write_f8821(data: dict):
    if SKIP_8821:
        return
    pdf = PyPDF2.PdfFileReader(F8821_PATH)
    writer = PyPDF2.PdfFileWriter()
    for page in pdf.pages:
        writer.addPage(page)
    address = f"{data['name']}{NEWLINE}"
    address += f"{data['address']}{NEWLINE}"
    address += f"{data['city']}, {data['state']} {data['zip']}"
    writer.update_page_form_field_values(
        writer.pages[0],
        {
            "f1_6[0]": address,
            "f1_7[0]": data["ein"],
            "f1_8[0]": data["phone"],
        },
    )
    filename = f"{data['name']} f8821.pdf"
    output_file = f"{OUTPUT_PATH}\\{filename}"
    writer.write(output_file)


def make_941x_dir():
    try:
        os.mkdir(OUTPUT_PATH)
    except FileExistsError:
        return  # Already exists, do nothing


def fix_zip(data):
    if data["company"]["zip"] is int:
        if data["company"]["zip"] < 10000:
            data["company"]["zip"] = f"0{data['company']['zip']}"


if __name__ == "__main__":
    pdf_reader = PyPDF2.PdfFileReader(F941X_PATH)
    pdf_writer = PyPDF2.PdfFileWriter()
    excel_wb = pd.ExcelFile(WS_PATH)
    data = {
        "company": extract_company_data(excel_wb.parse(sheet_name=SHEETS["input"])),
        "2020_q2": extract_tax_data(
            excel_wb.parse(sheet_name=SHEETS["2020Q2"]), ROW_2020
        ),
        "2020_q3": extract_tax_data(
            excel_wb.parse(sheet_name=SHEETS["2020Q3"]), ROW_2020
        ),
        "2020_q4": extract_tax_data(
            excel_wb.parse(sheet_name=SHEETS["2020Q4"]), ROW_2020
        ),
        "2021_q1": extract_tax_data(
            excel_wb.parse(sheet_name=SHEETS["2021Q1"]), ROW_2021
        ),
        "2021_q2": extract_tax_data(
            excel_wb.parse(sheet_name=SHEETS["2021Q2"]), ROW_2021
        ),
        "2021_q3": extract_tax_data(
            excel_wb.parse(sheet_name=SHEETS["2021Q3"]), ROW_2021
        ),
    }
    fix_zip(data)
    make_941x_dir()
    write_f8821(data["company"])
    for i in range(0, 6):
        pdf_writer.add_page(pdf_reader.pages[i])
    for year, quarter in YEAR_QUARTER:
        write_pdf_data(pdf_writer, data, year, quarter)
        update_quater_check_box(pdf_writer.pages[0], quarter)
        filename = f"{data['company']['name']} f941x {year} Q{quarter}.pdf"
        output_file = f"{OUTPUT_PATH}\\{filename}"
        pdf_writer.write(output_file)
