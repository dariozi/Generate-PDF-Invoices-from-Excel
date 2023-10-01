import pandas as pd
import glob
#To open excel file we need to install openpyxl from Python Packages

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath,sheet_name="Sheet 1")
    print(df)
