# In order to read 'xlsx' files there is a need to import pandas and install the 'openpyxl' package.
import pandas as pd
import glob
from fpdf import FPDF
# Library for extracting the file's name without it's start.
from pathlib import Path


# Returns a list with all the files in the 'invoices' directory that ends with '.xlsx'.
filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    # Creating the pdf object.
    pdf = FPDF(orientation="portrait", unit="mm", format="A4")

    # Extracting the name of the file without it's start and extension by using the 'stem()' func.
    filename = Path(filepath).stem
    # Extracting the invoice number and the date separately. (The 'split()' func returns a list).
    invoices_nr, date = filename.split("-")

    # Creating the pdf page.
    pdf.add_page()
    # Displaying the invoice number, on the PDF file.
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Invoices nr.{invoices_nr}", ln=1)
    # Displaying the date of that invoice, on the PDF file.
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    # Reading the 'xlsx' files content. The 'df' variable contains now all the content in a table.
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Extracting all the headers of the table. We convert it to a 'list' because 'df.columns' returns
    # all the columns as an 'index' type.
    header_columns = list(df.columns)
    # Replacing and titling all the columns headers.
    header_columns = [item.replace("_", " ").title() for item in header_columns]
    # Displaying all the columns headers in a table on the pdf.
    pdf.set_font(family="Times", style="B", size=10)
    # Setting the color to grey.
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=header_columns[0], border=1)
    pdf.cell(w=70, h=8, txt=header_columns[1], border=1)
    pdf.cell(w=30, h=8, txt=header_columns[2], border=1)
    pdf.cell(w=30, h=8, txt=header_columns[3], border=1)
    pdf.cell(w=30, h=8, txt=header_columns[4], border=1, ln=1)

    # Extracting all the rest of the table's content by itirating over each row in the 'df' table variable.
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        # Setting the color to grey.
        pdf.set_text_color(80, 80, 80)
        # The 'row' variable returns an 'int' so we convert it to 'str'.
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Calculating the total price and displaying at the bottom of the table on the ODF file.
    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    # Displaying the total sum sentence on the PDF file.
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=f"The total price is {total_sum}", ln=1)

    # Displaying the company name and logo on the PDF file.
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=25, h=8, txt=f"Sheet2Receipt")

    # Storing the pdf file that we just generated, in the 'PDFs' directory.
    pdf.output(f"PDFs/{filename}.pdf")
