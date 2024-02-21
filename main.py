from fpdf import FPDF
import pandas as pd
import glob
from pathlib import Path

# Filepath of the Excel invoices, can also be absolute path
FILEPATH_SOURCE = "sample-invoices/*.xlsx"
FILEPATH_DESTINATION = "pdf-invoices"

# Creating a list of filepaths of all the files
filepaths = glob.glob(FILEPATH_SOURCE)
print(filepaths)

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    filename = Path(filepath).stem
    print(filename)
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_font(family="Times", size=18, style="B")
    pdf.cell(0, 0, txt=f"Invoice No.: {filename.split('-')[0]}", ln=1)
    pdf.output(f"{FILEPATH_DESTINATION}/{filename}.pdf")
