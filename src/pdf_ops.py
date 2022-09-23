import PyPDF2
from PyPDF2.generic import NameObject
from pprint import pprint as pp
import datetime
import numpy as np
import os

NEWLINE = os.linesep


class pdf_helper:
    def __init__(
        self,
        template_path: str,
        f8821_path: str,
        write_path: str,
        quarter_fields: dict,
        pdf_dict: dict,
        skip_8821: bool,
        data: dict,
    ):
        self.reader = PyPDF2.PdfFileReader(template_path)
        self.load_writer()
        self.f8821_path = f8821_path
        self.write_path = write_path
        self.quarter_fields = quarter_fields
        self.pdf_dict = pdf_dict
        self.skip_8821 = skip_8821
        self.data = data
        self.make_941x_dir()
        self.write_f8821()

    def load_writer(self):
        """Load the writer and populate it with the pages read"""
        self.writer = PyPDF2.PdfFileWriter()
        for page in self.reader.pages:
            self.writer.addPage(page)

    def print_non_empy_fields(self):
        """Print the non-empty fields in the PDF.  This is used to identify keys of
        forms that we care about.
        """
        field_dict = self.reader.getFormTextFields()
        for k, v in field_dict.items():
            if v:
                print(f"key:{k}, value:{v}")

    def get_page_object_data(self, page: int = 0):
        """Print page objects.  The main goal here is to get the keys for page data

        Args:
            page (int): Which page to look at, we really only care about the first page
        """
        page = self.reader.pages[page]
        for i in range(0, len(page["/Annots"])):
            tmp = page["/Annots"][i].get_object()
            if "c1_" in tmp.get("/T"):
                pp(tmp)

    def update_quater_check_box(self, quarter: int):
        """Update the quarter checkbox on page 1

        Args:
            quarter (int): The quarter to set
        """
        page = self.reader.pages[0]
        for i in range(0, len(page["/Annots"])):
            annot = page["/Annots"][i].getObject()
            for k, field in self.quarter_fields.items():
                if annot.get("/T") == field:
                    as_value = "/Off"
                    if k == quarter:
                        as_value = f"/{quarter}"
                    annot.update(
                        {
                            NameObject("/V"): NameObject(as_value),
                            NameObject("/AS"): NameObject(as_value),
                        }
                    )

    def update_pdf_data(self, year: int, quarter: int):
        """This has the job of updating the PDF template for a specific year/quarter

        Args:
            year (int): the year  (2020|2021)
            quarter (int): The quarter (1|2|3|4)
        """
        d = datetime.date.today()
        dollars_18a, cents_18a = self.extract_dollars_and_cents(
            self.data[f"{year}_q{quarter}"]["18a"]
        )
        dollars_26a, cents_26a = self.extract_dollars_and_cents(
            self.data[f"{year}_q{quarter}"]["26a"]
        )
        dollars_27, cents_27 = self.extract_dollars_and_cents(
            self.data[f"{year}_q{quarter}"]["27"]
        )
        dollars_30, cents_30 = self.extract_dollars_and_cents(
            self.data[f"{year}_q{quarter}"]["30"]
        )
        self.writer.update_page_form_field_values(
            self.writer.pages[0],
            {
                self.pdf_dict["p1"]["ein1"]: self.data["company"]["ein"][:2],
                self.pdf_dict["p1"]["ein2"]: self.data["company"]["ein"][3:],
                self.pdf_dict["p1"]["name"]: self.data["company"]["name"],
                self.pdf_dict["p1"]["trade name"]: self.data["company"]["trade name"],
                self.pdf_dict["p1"]["address"]: self.data["company"]["address"],
                self.pdf_dict["p1"]["city"]: self.data["company"]["city"],
                self.pdf_dict["p1"]["state"]: self.data["company"]["state"],
                self.pdf_dict["p1"]["zip"]: self.data["company"]["zip"],
                self.pdf_dict["p1"]["city"]: self.data["company"]["city"],
                self.pdf_dict["p1"]["yyyy"]: d.year,
                self.pdf_dict["p1"]["dd"]: f"{d:%d}",
                self.pdf_dict["p1"]["mm"]: f"{d:%m}",
                self.pdf_dict["p1"]["year correcting"]: year,
            },
        )
        self.writer.update_page_form_field_values(
            self.writer.pages[1],
            {
                self.pdf_dict["p2"]["ein"]: self.data["company"]["ein"],
                self.pdf_dict["p2"]["name"]: self.data["company"]["name"],
                self.pdf_dict["p2"]["year"]: year,
                self.pdf_dict["p2"]["quarter"]: quarter,
                self.pdf_dict["tax"]["18a_dollars"][0]: dollars_18a,
                self.pdf_dict["tax"]["18a_dollars"][1]: dollars_18a,
                self.pdf_dict["tax"]["18a_dollars"][2]: -1 * dollars_18a,
                self.pdf_dict["tax"]["18a_cents"][0]: cents_18a,
                self.pdf_dict["tax"]["18a_cents"][1]: cents_18a,
                self.pdf_dict["tax"]["18a_cents"][2]: cents_18a,
            },
        )
        self.writer.update_page_form_field_values(
            self.writer.pages[2],
            {
                self.pdf_dict["p3"]["ein"]: self.data["company"]["ein"],
                self.pdf_dict["p3"]["name"]: self.data["company"]["name"],
                self.pdf_dict["p3"]["year"]: year,
                self.pdf_dict["p3"]["quarter"]: quarter,
                self.pdf_dict["tax"]["23_dollars"]: -1 * dollars_18a,
                self.pdf_dict["tax"]["23_cents"]: cents_18a,
                self.pdf_dict["tax"]["26a_dollars"][0]: dollars_26a,
                self.pdf_dict["tax"]["26a_dollars"][1]: dollars_26a,
                self.pdf_dict["tax"]["26a_dollars"][2]: -1 * dollars_26a,
                self.pdf_dict["tax"]["26a_cents"][0]: cents_26a,
                self.pdf_dict["tax"]["26a_cents"][1]: cents_26a,
                self.pdf_dict["tax"]["26a_cents"][2]: cents_26a,
                self.pdf_dict["tax"]["27_dollars"]: -1 * dollars_27,
                self.pdf_dict["tax"]["27_cents"]: cents_27,
                self.pdf_dict["tax"]["30_dollars"][0]: dollars_30,
                self.pdf_dict["tax"]["30_dollars"][1]: dollars_30,
                self.pdf_dict["tax"]["30_cents"][0]: cents_30,
                self.pdf_dict["tax"]["30_cents"][1]: cents_30,
            },
        )
        self.writer.update_page_form_field_values(
            self.writer.pages[3],
            {
                self.pdf_dict["p4"]["ein"]: self.data["company"]["ein"],
                self.pdf_dict["p4"]["name"]: self.data["company"]["name"],
                self.pdf_dict["p4"]["year"]: year,
                self.pdf_dict["p4"]["quarter"]: quarter,
            },
        )
        self.writer.update_page_form_field_values(
            self.writer.pages[4],
            {
                self.pdf_dict["p5"]["ein"]: self.data["company"]["ein"],
                self.pdf_dict["p5"]["name"]: self.data["company"]["name"],
                self.pdf_dict["p5"]["year"]: year,
                self.pdf_dict["p5"]["quarter"]: quarter,
            },
        )

    def extract_dollars_and_cents(self, num: np.float64) -> list:
        if int(np.round(num)) == -1 or int(np.round(num)) == 0:
            return ["", ""]
        dollars = int(np.floor(num))
        # dollars = add_commas_to_dollars(dollars)
        cents = str(int(num * 100))[-2:]  # don't forget the colon
        return [dollars, cents]

    def write_pdf_file(self, year: int, quarter: int):
        filename = (
            f"{self.data['company']['name']} f941x {year} Q{quarter}.pdf"
        )
        output_file = f"{self.write_path}\\{filename}"
        self.writer.write(output_file)

    def write_f8821(self):
        if self.skip_8821:
            return
        data = self.data["company"]
        pdf = PyPDF2.PdfFileReader(self.f8821_path)
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
        output_file = f"{self.write_path}\\{filename}"
        writer.write(output_file)

    def make_941x_dir(self):
        try:
            os.mkdir(self.write_path)
        except FileExistsError:
            return  # Already exists, do nothing

    def make_pdf(self, year, quarter):
        self.update_pdf_data(year, quarter)
        self.update_quater_check_box(quarter)
        self.write_pdf_file(year, quarter)
