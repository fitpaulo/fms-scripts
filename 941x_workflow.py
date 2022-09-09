import yaml
import src.pdf_ops as pdf_ops
import src.excel_ops as excel_ops

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


def fix_zip(data):
    if data["company"]["zip"] is int:
        if data["company"]["zip"] < 10000:
            data["company"]["zip"] = f"0{data['company']['zip']}"


if __name__ == "__main__":
    pdf_reader = pdf_ops.create_pdf_reader(F941X_PATH)
    pdf_writer = pdf_ops.create_pdf_writer()
    wb = excel_ops.load_wb(WB_PATH)
    data = excel_ops.load_data(wb, SHEETS, ROW_2020, ROW_2021, ROUND_DELTA)
    fix_zip(data)
    pdf_ops.make_941x_dir()
    pdf_ops.write_f8821(data["company"], F8821_PATH, OUTPUT_PATH, SKIP_8821)
    for i in range(0, 6):
        pdf_writer.add_page(pdf_reader.pages[i])
    for year, quarter in YEAR_QUARTER:
        pdf_ops.write_pdf_data(pdf_writer, data, year, quarter, PDF_DICT)
        pdf_ops.update_quater_check_box(pdf_writer.pages[0], quarter, QUARTER_FIELDS)
        filename = f"{data['company']['name']} f941x {year} Q{quarter}.pdf"
        output_file = f"{OUTPUT_PATH}\\{filename}"
        pdf_ops.write(pdf_writer, output_file)
