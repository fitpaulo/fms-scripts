import PyPDF2
from PyPDF2.generic import NameObject
import datetime
import numpy as np
import os


NEWLINE = os.linesep


def print_non_empty_fields(pdf: PyPDF2.PdfFileReader):
    field_dict = pdf.getFormTextFields()
    for k, v in field_dict.items():
        if v:
            print(f"key:{k}   value:{v}")


def update_quater_check_box(page: PyPDF2._page.PageObject, value: int, quarter_fields: dict):
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
            # from pprint import pprint as pp
            # pp(tmp)
            pass


def extract_dollars_and_cents(num: np.float63) -> list:
    if int(np.round(num)) == -1:
        return ["", ""]
    dollars = int(np.floor(num))
    # dollars = add_commas_to_dollars(dollars)
    cents = str(num)[-3:]  # Don't forget the colon
    if cents[-1] == ".":
        cents = f"{cents[0]}0"
    return [dollars, cents]


def write_pdf_data(pdf: PyPDF2.PdfFileWriter, data: dict, year: int, quarter: int, pdf_dict: dict):
    d = datetime.date.today()
    dollars_18a, cents_18a = extract_dollars_and_cents(
        data[f"{year}_q{quarter}"]["18a"]
    )
    dollars_26a, cents_26a = extract_dollars_and_cents(
        data[f"{year}_q{quarter}"]["26a"]
    )
    dollars_27, cents_27 = extract_dollars_and_cents(data[f"{year}_q{quarter}"]["27"])
    dollars_30, cents_30 = extract_dollars_and_cents(data[f"{year}_q{quarter}"]["30"])
    pdf.update_page_form_field_values(
        pdf.pages[0],
        {
            pdf_dict["p1"]["ein1"]: data["company"]["ein"][:2],
            pdf_dict["p1"]["ein2"]: data["company"]["ein"][3:],
            pdf_dict["p1"]["name"]: data["company"]["name"],
            pdf_dict["p1"]["trade name"]: data["company"]["trade name"],
            pdf_dict["p1"]["address"]: data["company"]["address"],
            pdf_dict["p1"]["city"]: data["company"]["city"],
            pdf_dict["p1"]["state"]: data["company"]["state"],
            pdf_dict["p1"]["zip"]: data["company"]["zip"],
            pdf_dict["p1"]["city"]: data["company"]["city"],
            pdf_dict["p1"]["yyyy"]: d.year,
            pdf_dict["p1"]["dd"]: f"{d:%d}",
            pdf_dict["p1"]["mm"]: f"{d:%m}",
            pdf_dict["p1"]["year correcting"]: year,
        },
    )
    pdf.update_page_form_field_values(
        pdf.pages[1],
        {
            pdf_dict["p2"]["ein"]: data["company"]["ein"],
            pdf_dict["p2"]["name"]: data["company"]["name"],
            pdf_dict["p2"]["year"]: year,
            pdf_dict["p2"]["quarter"]: quarter,
            pdf_dict["tax"]["18a_dollars"][0]: dollars_18a,
            pdf_dict["tax"]["18a_dollars"][1]: dollars_18a,
            pdf_dict["tax"]["18a_dollars"][2]: -1 * dollars_18a,
            pdf_dict["tax"]["18a_cents"][0]: cents_18a,
            pdf_dict["tax"]["18a_cents"][1]: cents_18a,
            pdf_dict["tax"]["18a_cents"][2]: cents_18a,
        },
    )
    pdf.update_page_form_field_values(
        pdf.pages[2],
        {
            pdf_dict["p3"]["ein"]: data["company"]["ein"],
            pdf_dict["p3"]["name"]: data["company"]["name"],
            pdf_dict["p3"]["year"]: year,
            pdf_dict["p3"]["quarter"]: quarter,
            pdf_dict["tax"]["23_dollars"]: -1 * dollars_18a,
            pdf_dict["tax"]["23_cents"]: cents_18a,
            pdf_dict["tax"]["26a_dollars"][0]: dollars_26a,
            pdf_dict["tax"]["26a_dollars"][1]: dollars_26a,
            pdf_dict["tax"]["26a_dollars"][2]: -1 * dollars_26a,
            pdf_dict["tax"]["26a_cents"][0]: cents_26a,
            pdf_dict["tax"]["26a_cents"][1]: cents_26a,
            pdf_dict["tax"]["26a_cents"][2]: cents_26a,
            pdf_dict["tax"]["27_dollars"]: -1 * dollars_27,
            pdf_dict["tax"]["27_cents"]: cents_27,
            pdf_dict["tax"]["30_dollars"][0]: dollars_30,
            pdf_dict["tax"]["30_dollars"][1]: dollars_30,
            pdf_dict["tax"]["30_cents"][0]: cents_30,
            pdf_dict["tax"]["30_cents"][1]: cents_30,
        },
    )
    pdf.update_page_form_field_values(
        pdf.pages[3],
        {
            pdf_dict["p4"]["ein"]: data["company"]["ein"],
            pdf_dict["p4"]["name"]: data["company"]["name"],
            pdf_dict["p4"]["year"]: year,
            pdf_dict["p4"]["quarter"]: quarter,
        },
    )
    pdf.update_page_form_field_values(
        pdf.pages[4],
        {
            pdf_dict["p5"]["ein"]: data["company"]["ein"],
            pdf_dict["p5"]["name"]: data["company"]["name"],
            pdf_dict["p5"]["year"]: year,
            pdf_dict["p5"]["quarter"]: quarter,
        },
    )


def write_f8821(data: dict, f8821_path: str, write_path, skip=False):
    if skip:
        return
    pdf = PyPDF2.PdfFileReader(f8821_path)
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
    output_file = f"{write_path}\\{filename}"
    writer.write(output_file)


def make_941x_dir(dir_path: str):
    try:
        os.mkdir(dir_path)
    except FileExistsError:
        return  # Already exists, do nothing


def create_pdf_reader(reader_path: str):
    return PyPDF2.PdfFileReader(reader_path)


def create_pdf_writer():
    return PyPDF2.PdfFileWriter()


def write_pdf_file(pdf: PyPDF2.PdfFileWriter, file_path: str):
    pdf.write(file_path)
