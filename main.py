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
    filename = Path(filepath).stem
    print(filename)

    invoice_no, date = filename.split("-")
    print(invoice_no, date)

    # Creating a PDF
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_font(family="Times", size=18, style="B")
    pdf.cell(0, 10, txt=f"Invoice No.: {invoice_no}", ln=1)
    pdf.cell(0, 10, txt=f"Invoice Date: {date}", ln=1)
    # Line spacer
    pdf.cell(0, 10, ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Getting headers from the Excel file in form of list
    headers = list(df.columns)
    # This list to match column width of the header row
    column_width = [30, 70, 35, 30, 30]
    flag = 0
    for header in headers:
        # Visual formatting of the header
        header = header.replace("_", " ")
        header = header.title()

        # This if-else condition is to provide line break for the last column
        if flag < 4:
            pdf.set_font(family="Times", size=10, style="B")
            pdf.cell(column_width[flag], 20, txt=header, border=1, align="C")
            flag += 1
        else:
            pdf.set_font(family="Times", size=10, style="B")
            pdf.cell(column_width[flag], 20, txt=header, border=1, ln=1, align="C")

    sub_total = 0.0
    # This loop is to insert the items from each row in the table
    for index, row in df.iterrows():
        print(index)
        pdf.set_font(family="Times", size=10)
        pdf.cell(30, 10, txt=str(row["product_id"]), border=1)
        pdf.cell(70, 10, txt=row["product_name"], border=1)
        pdf.cell(35, 10, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(30, 10, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(30, 10, txt=str(row["total_price"]), border=1, ln=1)
        sub_total += float(row["total_price"])
    # Column spacer to set the total amount column
    pdf.cell(135, 10)
    # Sub Total columns
    pdf.cell(30, 10, txt="Sub Total", border=1, align="C")
    pdf.cell(30, 10, txt=str(sub_total), border=1, ln=1)

    # Column spacer to set the total amount column
    pdf.cell(135, 10)
    # VAT columns
    pdf.cell(30, 10, txt="VAT", border=1, align="C")
    vat_amount = sub_total * 0.05
    pdf.cell(30, 10, txt=str(vat_amount), border=1, ln=1)

    # Column spacer to set the total amount column
    pdf.cell(135, 10)
    # Grand Total columns
    pdf.cell(30, 10, txt="Grand Total", border=1, align="C")
    grand_total = sub_total + vat_amount
    pdf.cell(30, 10, txt=str(grand_total), border=1, ln=1)

    # Line spacer
    pdf.cell(0, 10, ln=1)
    # Footer
    pdf.set_font(family="Times", size=12, style="B")
    pdf.cell(0, 10, txt=f"The total amount due is {grand_total} Euros.", ln=1)
    # Company details
    pdf.cell(0, 12, txt=f"Codesabers Technology")
    # pdf.image("IMAGE_LINK", w=, h=)

    # Generating the final pdf file
    pdf.output(f"{FILEPATH_DESTINATION}/{filename}.pdf")
