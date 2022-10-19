import os
import shutil
import sys

import yaml

from src.excel_ops import excel_helper
from src.pdf_ops import pdf_helper

with open("conf/config.yaml", "r") as file:
    conf = yaml.safe_load(file)
with open("conf/f941x.yaml", "r") as file:
    pdf_conf = yaml.safe_load(file)

# From yaml
COMPANY_TYPE = conf["type"]
SKIP_8821 = conf["skip"]
YEAR_QUARTER = conf["year_quarter"]
DROPBOX_PATH = conf["dropbox_path"]
PDF_DICT = pdf_conf["pdf_dict"]
QUARTER_FIELDS = pdf_conf["quarter_fields"]
SHEETS = conf["excel_sheet_names"]
ROUND_DELTA = conf["round_delta"]
ROW_2020 = conf["row_2020"]
ROW_2021 = conf["row_2021"]
BASE_PATH = os.path.join(DROPBOX_PATH, f"COMPANIES {COMPANY_TYPE}")
PDF_PATH = os.path.join(BASE_PATH, f"PAT {COMPANY_TYPE} ERTC")
F941X_PATH = os.path.join(PDF_PATH, conf["f941x_file_name"])
F8821_PATH = os.path.join(PDF_PATH, conf["f8821_file_name"])


def build_company_path():
    company_path = os.path.join(BASE_PATH, conf["company"][0])
    if os.path.exists(company_path):
        contents = os.listdir(company_path)
        for company in contents:
            if conf["company"] in company:
                company_path = os.path.join(company_path, company)
                break
        contents = os.listdir(company_path)
        if len(contents) == 1:
            company_path = os.path.join(company_path, contents[0])
        return company_path
    else:
        raise RuntimeError(f"Company path does not exist: {company_path}")


def copy_worksheet(dest_dir: str):
    for item in os.listdir(PDF_PATH):
        if conf["base_ws_name"] in item:
            shutil.copy(
                os.path.join(PDF_PATH, item),
                dest_dir,
            )
            new_name = f"{conf['company']} ERTC Worksheet.xlsx"
            os.rename(
                os.path.join(dest_dir, item),
                os.path.join(dest_dir, new_name)
            )


def set_up_worksheet():
    new_dir_name = os.path.join(COMPANY_PATH, "Payroll And Worksheet")
    os.rename(
        os.path.join(COMPANY_PATH, "Payroll"),
        new_dir_name,
    )
    copy_worksheet(new_dir_name)


def build_wb_path():
    if os.path.exists(os.path.join(COMPANY_PATH, "Payroll")):
        # Then set up the file and exit
        set_up_worksheet()
        sys.exit(0)
    for dir in conf["payroll_dirs"]:
        if os.path.exists(os.path.join(COMPANY_PATH, dir)):
            wb_path = os.path.join(COMPANY_PATH, dir)
            contents = os.listdir(wb_path)
            for file in contents:
                if "worksheet" in file.lower():
                    return os.path.join(wb_path, file)
            raise RuntimeError(f"Cannot find the WS in: {wb_path}")
    raise RuntimeError(
        f"Cannot find the 'Payroll and Worksheet' dir in: {COMPANY_PATH}"
    )


def validate_path(path):
    if os.path.exists(path):
        return True
    raise RuntimeError(f"Path does not exist: {path}")


# These var requie some functions to build
COMPANY_PATH = build_company_path()
WB_PATH = build_wb_path()
OUTPUT_PATH = os.path.join(COMPANY_PATH, "941x")

if __name__ == "__main__":
    validate_path(F8821_PATH)
    validate_path(F941X_PATH)
    excel = excel_helper(WB_PATH, SHEETS, ROUND_DELTA, ROW_2020, ROW_2021)
    excel.load_data()
    pdf = pdf_helper(
        F941X_PATH,
        F8821_PATH,
        OUTPUT_PATH,
        QUARTER_FIELDS,
        PDF_DICT,
        SKIP_8821,
        excel.data,
    )
    for year, quarter in YEAR_QUARTER:
        pdf.make_pdf(year, quarter)
