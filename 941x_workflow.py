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


def build_company_path(company: str):
    if company[0] in "1234567890":
        company_path = os.path.join(BASE_PATH, "1234567890")
    else:
        company_path = os.path.join(BASE_PATH, company[0])
    for item in os.listdir(company_path):
        if company.lower() in item.lower():
            company_path = os.path.join(company_path, company)
            break
    if len(os.listdir(company_path)) == 1:
        company_path = os.path.join(company_path, os.listdir(company_path)[0])
    return company_path


def copy_worksheet(dest_dir: str):
    for item in os.listdir(PDF_PATH):
        if conf["base_ws_name"] in item:
            shutil.copy(
                os.path.join(PDF_PATH, item),
                dest_dir,
            )
            new_name = f"{conf['company']} ERTC Worksheet.xlsx"
            os.rename(os.path.join(dest_dir, item), os.path.join(dest_dir, new_name))


def set_up_worksheet(company_path: str):
    new_dir_name = os.path.join(company_path, "Payroll And Worksheet")
    os.rename(
        os.path.join(company_path, "Payroll"),
        new_dir_name,
    )
    copy_worksheet(new_dir_name)


def build_wb_path(company_path: str):
    if os.path.exists(os.path.join(company_path, "Payroll")):
        # Then set up the file and exit
        set_up_worksheet(company_path)
        sys.exit(0)
    for item in os.listdir(company_path):
        if item.lower() in conf["payroll_dirs"]:
            wb_path = os.path.join(company_path, item)
            contents = os.listdir(wb_path)
            for file in contents:
                if "worksheet" in file.lower():
                    return os.path.join(wb_path, file)
            raise RuntimeError(f"Cannot find the WS in: {wb_path}")
    raise RuntimeError(
        f"Cannot find the 'Payroll and Worksheet' dir in: {company_path}"
    )


def validate_path(path):
    if os.path.exists(path):
        return True
    raise RuntimeError(f"Path does not exist: {path}")


if __name__ == "__main__":
    validate_path(F8821_PATH)
    validate_path(F941X_PATH)
    errors = []
    error_companies = []
    for company in conf["companies"]:
        try:
            company_path = build_company_path(company)
            wb_path = build_wb_path(company_path)
            output_path = os.path.join(company_path, "941x")
            excel = excel_helper(wb_path, SHEETS, ROUND_DELTA, ROW_2020, ROW_2021)
            excel.load_data()
            pdf = pdf_helper(
                F941X_PATH,
                F8821_PATH,
                output_path,
                QUARTER_FIELDS,
                PDF_DICT,
                SKIP_8821,
                excel.data,
            )
            for year, quarter in YEAR_QUARTER:
                pdf.make_pdf(year, quarter)
        except Exception as e:
            errors.append(e)
            error_companies.append(company)
    if errors:
        print(errors)
        print(error_companies)
