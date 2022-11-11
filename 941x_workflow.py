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
SKIP_8821 = conf["skip"]
YEAR_QUARTER = conf["year_quarter"]
DROPBOX_PATH = conf["dropbox_path"]
PDF_DICT = pdf_conf["pdf_dict"]
QUARTER_FIELDS = pdf_conf["quarter_fields"]
SHEETS = conf["excel_sheet_names"]
ROUND_DELTA = conf["round_delta"]
ROW_2020 = conf["row_2020"]
ROW_2021 = conf["row_2021"]


def update_company_path(company_path: str):
    contents = os.listdir(company_path)
    if len(contents) == 1:
        return os.path.join(company_path, contents[0])
    return company_path


def get_company_paths(company: str):
    res = []
    for company_type in conf["types"]:
        base_path = os.path.join(DROPBOX_PATH, f"COMPANIES {company_type}")
        if company[0] in "1234567890":
            company_path = os.path.join(base_path, "1234567890")
        else:
            company_path = os.path.join(base_path, company[0])
        for item in os.listdir(company_path):
            if company in item:
                company_path = update_company_path(os.path.join(company_path, item))
                res.append(os.path.join(base_path, f"PAT {company_type} ERTC"))
                res.append(company_path)
    if len(res) == 0:
        raise RuntimeError(f"Unable to find ${company} under TSP or LA")
    return res


def copy_worksheet(dest_dir: str, pdf_path: str, company: str):
    for item in os.listdir(pdf_path):
        if conf["base_ws_name"] in item:
            shutil.copy(
                os.path.join(pdf_path, item),
                dest_dir,
            )
            new_name = f"{company} ERTC Worksheet.xlsx"
            os.rename(os.path.join(dest_dir, item), os.path.join(dest_dir, new_name))
            break


def set_up_worksheet(company_path: str, pdf_path: str, company: str):
    new_dir_name = os.path.join(company_path, "Payroll And Worksheet")
    os.rename(
        os.path.join(company_path, "Payroll"),
        new_dir_name,
    )
    copy_worksheet(new_dir_name, pdf_path, company)


def build_wb_path(company_path: str, pdf_path: str, company: str):
    if os.path.exists(os.path.join(company_path, "Payroll")):
        # Then set up the file and exit
        set_up_worksheet(company_path, pdf_path, company)
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
    errors = []
    error_companies = []
    for company in conf["companies"]:
        try:
            pdf_path, company_path = get_company_paths(company)
            f941x_path = os.path.join(pdf_path, conf["f941x_file_name"])
            f8821_path = os.path.join(pdf_path, conf["f8821_file_name"])
            output_path = os.path.join(company_path, "941x")
            wb_path = build_wb_path(company_path, pdf_path, company)
            excel = excel_helper(wb_path, SHEETS, ROUND_DELTA, ROW_2020, ROW_2021)
            excel.load_data()
            pdf = pdf_helper(
                f941x_path,
                f8821_path,
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
