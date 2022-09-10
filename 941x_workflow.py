import yaml
from src.excel_ops import excel_helper
from src.pdf_ops import pdf_helper

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
ROUND_DELTA = conf["round_delta"]
ROW_2020 = conf["row_2020"]
ROW_2021 = conf["row_2021"]

# Dynamic vars
BASE_PATH = f"{DROPBOX_PATH}\\COMPANIES {COPANY_TYPE}"
PDF_PATH = f"{BASE_PATH}\\PAT {COPANY_TYPE} ERTC"
F941X_PATH = f"{PDF_PATH}\\f941x 8-9-22.pdf"
F8821_PATH = f"{PDF_PATH}\\f8821 8-9-22.pdf"
WB_PATH = f"{BASE_PATH}\\{COMPANY_PATH}\\{PAYROLL_DIR}\\{WS_NAME}.xlsx"
OUTPUT_PATH = f"{BASE_PATH}\\{COMPANY_PATH}\\941x"


if __name__ == "__main__":
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
