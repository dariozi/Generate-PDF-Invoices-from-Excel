# To manipulate Data
import pandas as pd
# Glob is to have all file in one filepath list
import glob
# To generate PDF
from fpdf import FPDF
# We use Path to extract the invoice number from the invoice name. Separate the full path and the file name
from pathlib import Path

# To open Excel file we need to install openpyxl from Python Packages

filepaths = glob.glob("invoices/*.xlsx")

# A For-Loop to eterate into each filepath
for filepath in filepaths:


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
    pdf.cell(w=50, h=8, txt=f"Date: {date}",ln=1)

    #Add 10mm of space
    pdf.ln(10)
    #Read the dataframe
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    #Creating the header
    columns = df.columns
    columns = [item.replace("_"," ").title() for item in columns]
    pdf.set_font(family="Times", style="B", size=10)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=60, h=8, txt=columns[1], border=1)
    pdf.cell(w=40, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    #inserting the excel line into the pdf. must be a string
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=60, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1,ln=1)

    #Create the output in PDF
    pdf.output(f"PDFs/{filename}.pdf")
