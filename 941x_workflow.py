import PyPDF2
from PyPDF2.generic import NameObject
import pandas as pd
import numpy as np
import datetime
import os


###########################################
#               EDIT THESE                #
###########################################
COMPANY_PATH = "I\\I.E.S KENTUCKY, LLC\\ERC"
COPANY_TYPE = "LA"
WS_NAME = "IES KENTUCKY ERTC Worksheet"  # WS = Worksheet
SKIP_8821 = False

# Comment out lines you don't want to make f941xs for
YEAR_QUARTER = [
    [2020, 2],
    [2020, 3],
    [2020, 4],
    [2021, 1],
    [2021, 2],
    [2021, 3],
]
##########################################
#            STOP EDITING                #
##########################################

BASE_PATH = f"C:\\Users\\dguim\\FMS Dropbox\\COMPANIES {COPANY_TYPE}"
PDF_PATH = f"{BASE_PATH}\\PAT {COPANY_TYPE} ERTC"
F941X_PATH = f"{PDF_PATH}\\f941x 8-9-22.pdf"
F8821_PATH = f"{PDF_PATH}\\f8821 8-9-22.pdf"
WS_PATH = f"{BASE_PATH}\\{COMPANY_PATH}\\Payroll And Worksheet\\{WS_NAME}.xlsx"
OUTPUT_PATH = f"{BASE_PATH}\\{COMPANY_PATH}\\941x"
NEWLINE = os.linesep


PDF_DICT = {
    "p1": {
        "ein1": "f1_01[0]",
        "ein2": "f1_02[0]",
        "name": "f1_03[0]",
        "trade name": "f1_04[0]",
        "address": "f1_05[0]",
        "city": "f1_06[0]",
        "state": "f1_07[0]",
        "zip": "f1_08[0]",
        "year correcting": "f1_12[0]",
        "mm": "f1_13[0]",
        "dd": "f1_14[0]",
        "yyyy": "f1_15[0]",
    },
    "p2": {
        "name": "f2_01[0]",
        "ein": "f2_02[0]",
        "quarter": "f2_03[0]",
        "year": "f2_04[0]",
    },
    "p3": {
        "name": "f3_01[0]",
        "ein": "f3_02[0]",
        "quarter": "f3_03[0]",
        "year": "f3_04[0]",
    },
    "p4": {
        "name": "f4_01[0]",
        "ein": "f4_02[0]",
        "quarter": "f4_03[0]",
        "year": "f4_04[0]",
    },
    "p5": {
        "name": "f5_01[0]",
        "ein": "f5_02[0]",
        "quarter": "f5_03[0]",
        "year": "f5_04[0]",
    },
    "tax": {
        "18a_dollars": [
            "f2_99[0]",
            "f2_103[0]",
            "f2_105[0]",  # negative dollars
        ],
        "18a_cents": [
            "f2_100[0]",
            "f2_104[0]",
            "f2_106[0]",
        ],
        "23_dollars": "f3_139[0]",  # negative
        "23_cents": "f3_140[0]",
        "26a_dollars": [
            "f3_157[0]",
            "f3_161[0]",
            "f3_163[0]",  # negative dollars
        ],
        "26a_cents": [
            "f3_158[0]",
            "f3_162[0]",
            "f3_164[0]",
        ],
        "27_dollars": "f3_181[0]",  # negative
        "27_cents": "f3_182[0]",
        "30_dollars": [
            "f3_195[0]",
            "f3_199[0]",
        ],
        "30_cents": [
            "f3_196[0]",
            "f3_200[0]",
        ],
    },
}


def print_non_empty_fields(pdf: PyPDF2.PdfFileReader):
    field_dict = pdf.getFormTextFields()
    for k, v in field_dict.items():
        if v:
            print(f"key:{k}   value:{v}")


def update_quater_check_box(page: PyPDF2._page.PageObject, value: int):
    quarter_fields = {1: "c1_02[0]", 2: "c1_02[1]", 3: "c1_02[2]", 4: "c1_02[3]"}
    for i in range(0, len(page["/Annots"])):
        annot = page["/Annots"][i].getObject()
        for k, field in quarter_fields.items():
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
            from pprint import pprint as pp

            pp(tmp)


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
        ](),
    }


def extract_tax_data(df: pd.DataFrame, row: int) -> dict:
    #  Note, rounding to 3 here to make round equal that in excel
    return {
        "18a": np.round(df.iloc[row, 3], decimals=3),
        "26a": np.round(df.iloc[row + 2, 3], decimals=3),
        "27": np.round(df.iloc[row + 4, 3], decimals=3),
        "30": np.round(df.iloc[row + 6, 3], decimals=3),
    }


def add_commas_to_dollars(num: int):
    step = 3
    out = []
    current = str(num)
    while len(current) > step:
        out.append(current[-1 * step:])
        current = current[:len(current) - step]
    out.append(current)
    out.reverse()
    return ",".join(out)


def extract_dollars_and_cents(num: np.float64) -> list:
    if int(np.round(num)) == 0:
        return ["", ""]
    num = np.round(num, decimals=2)
    dollars = int(np.floor(num))
    # dollars = add_commas_to_dollars(dollars)
    cents = str(num)[-2:]  # Don't forget the colon
    if cents[0] == ".":
        cents = f"{cents[1]}0"
    return [dollars, cents]


def write_pdf_data(pdf: PyPDF2.PdfFileWriter, data: dict, year: int, quarter: int):
    d = datetime.date.today()
    dollars_18a, cents_18a = extract_dollars_and_cents(
        data[f"q{quarter}_{year}"]["18a"]
    )
    dollars_26a, cents_26a = extract_dollars_and_cents(
        data[f"q{quarter}_{year}"]["26a"]
    )
    dollars_27, cents_27 = extract_dollars_and_cents(data[f"q{quarter}_{year}"]["27"])
    dollars_30, cents_30 = extract_dollars_and_cents(data[f"q{quarter}_{year}"]["30"])
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


if __name__ == "__main__":
    pdf_reader = PyPDF2.PdfFileReader(F941X_PATH)
    pdf_writer = PyPDF2.PdfFileWriter()
    excel_wb = pd.ExcelFile(WS_PATH)
    data = {
        "company": extract_company_data(excel_wb.parse(sheet_name="Data Input")),
        "q2_2020": extract_tax_data(excel_wb.parse(sheet_name="2020 Q2 941 Calc"), 46),
        "q3_2020": extract_tax_data(excel_wb.parse(sheet_name="2020 Q3 941 Calc"), 46),
        "q4_2020": extract_tax_data(excel_wb.parse(sheet_name="2020 Q4 941 Calc"), 46),
        "q1_2021": extract_tax_data(excel_wb.parse(sheet_name="2021 Q1 941 Calc"), 47),
        "q2_2021": extract_tax_data(excel_wb.parse(sheet_name="2021 Q2 941 Calc"), 47),
        "q3_2021": extract_tax_data(excel_wb.parse(sheet_name="2021 Q3 941 Calc"), 47),
    }
    write_f8821(data["company"])
    for i in range(0, 6):
        pdf_writer.add_page(pdf_reader.pages[i])
    for year, quarter in YEAR_QUARTER:
        write_pdf_data(pdf_writer, data, year, quarter)
        update_quater_check_box(pdf_writer.pages[0], quarter)
        filename = f"{data['company']['name']} f941x {year} Q{quarter}.pdf"
        output_file = f"{OUTPUT_PATH}\\{filename}"
        pdf_writer.write(output_file)
