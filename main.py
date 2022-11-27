# Fourth program - Creates PDF invoices out of Excel files.

# In order to read 'xlsx' files there is a need to import pandas and install the
# 'openpyxl' package.
import pandas as pd
import glob

# Returns a list with all the files in the 'invoices' directory that ends with '.xlsx'.
filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
