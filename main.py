# To manipulate Data
import pandas as pd
# Glob is to have all fine in one filepath list
import glob
# To generate PDF
from fpdf import FPDF
# We use Path to extract the invoice number from the invoice name
from pathlib import Path

# To open Excel file we need to install openpyxl from Python Packages

filepaths = glob.glob("invoices/*.xlsx")

# A For-Loop to eterate into each filepath
for filepath in filepaths:
    #Read the dataframe
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    #Create the PDF
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # this will remove the folder name giving us back the file name as a string
    filename = Path(filepath).stem

    # We get the string, we spil it where the - is and we get the first item of the list
    invoice_nr = filename.split("-")[0]
    date = filename.split("-")[1]

    #Font and size of the Cell
    pdf.set_font(family="Times", size=16, style="B")

    # Invoice Nr will be printed as the first cell of the PDF document
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}")

    #New line on the PDF file
    pdf.ln()

    #Inserting the date from the file name
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}")

    #Create the output in PDF
    pdf.output(f"PDFs/{filename}.pdf")
