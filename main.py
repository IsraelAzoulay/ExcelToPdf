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
    # Extracting the first part of the file's name.
    invoices_nr = filename.split("-")[0]

    pdf.add_page()
    pdf.set_font(family="Times", style="B", size=16)
    # Creating the header of the pdf file.
    pdf.cell(w=50, h=8,txt=f"Invoices nr.{invoices_nr}")
    # Storing the pdf file that we just generated, in the 'PDFs' directory.
    pdf.output(f"PDFs/{filename}.pdf")
