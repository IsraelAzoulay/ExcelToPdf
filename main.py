# Fourth program - Creates PDF invoices out of Excel files.

# In order to read 'xlsx' files there is a need to import pandas and install the 'openpyxl' package.
import pandas as pd
import glob
from fpdf import FPDF
# Library for extracting the file's name without it's start and extention.
from pathlib import Path


# Returns a list with all the files in the 'invoices' directory that ends with '.xlsx'.
filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="portrait", unit="mm", format="A4")

    # Extracting the name of the file without it's start and extention by using the 'stem()' func.
    filename = Path(filepath).stem
    # Extracting the invoice number and the date separately.
    invoices_nr, date = filename.split("-")

    pdf.add_page()
    # Displaying the invoice number, on the PDF file.
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8,txt=f"Invoices nr.{invoices_nr}", ln=1)
    # Displaying the date of that invoice, on the PDF file.
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8,txt=f"Invoices nr.{date}")



    # Storing the pdf file that we just generated, in the 'PDFs' directory.
    pdf.output(f"PDFs/{filename}.pdf")










txt_pdf = FPDF(orientation="portrait", unit="mm", format="A4")
txt_filepaths = glob.glob("TEXTs_Files/*.txt")
for filepath in txt_filepaths:
    filename = Path(filepath).stem
    name = filename.capitalize()

    txt_pdf.add_page()
    txt_pdf.set_font(family="Times", style="B", size=16)
    txt_pdf.cell(w=50, h=8, txt=name, ln=1)

txt_pdf.output("output.pdf")



