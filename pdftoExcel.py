import os
import PyPDF2
import openpyxl
import re


# Constants
PDF_REGEX = re.compile(r'[^\\]*\.pdf')
EXCEL_HEADER = ["JOB DOCKET NO.", "JOB DOCKET DATE", "JOB NAME", "ORD.QTY", "PRINT QTY", "PRINT QTY WITH WASTAGE", "DIMENSIONS", "Prod. Type", "Collar Flap", "Side Paste", "Tuck Flap", "Order Qty", "Extra Qty", "Print Qty", "Wastage Qty", "Final Job Qty", "Board Details", "UPS", "SHEETS", "KG", "CUT SIZE", "CUT SHEET", "WASTAGE SHEET", "CUTS", "CUT SHEET UPS", "TOTAL CUT SHEETS", "WASTAGE"]

# Creates a list of PDF file names
pdfFileList = []
FileList = os.listdir(os.getcwd())
for file in FileList:
    if (PDF_REGEX.search(file)):
        pdfFileList.append(file)

data = []

for pdf in pdfFileList:
    try:
        with open(pdf, 'rb') as pdfFileObj:
            pdfReader = PyPDF2.PdfReader(pdfFileObj)
            pageObj = pdfReader.pages[0]
            extractedData = pageObj.extract_text()

            jobDocketNoExtracRegex = re.compile(r'(JC\/\d\d-\d\d\/\d\d\d\d)')
            jobDockNoMatch = jobDocketNoExtracRegex.search(extractedData)

            jobDocketDateExtracRegex = re.compile(r'JOB DOCKET DATE (\d{2}\/\d{2}\/\d{4})')
            jobDocketDateMatch = jobDocketDateExtracRegex.search(extractedData)

            if jobDockNoMatch and jobDocketDateMatch:
                jobDockNo = jobDockNoMatch.group()
                jobDocketDate = jobDocketDateMatch.group(1)
                data.append([jobDockNo, jobDocketDate])
            else:
                print(f"No match found in {pdf}")
    except Exception as e:
        print(f"Error processing {pdf}: {e}")

# Create Excel File
wb = openpyxl.Workbook()
sheet = wb.active

# Insert header row
sheet.append(EXCEL_HEADER)

# Insert data rows
for row in data:
    sheet.append(row)

# Save Excel file
wb.save('dataFile.xlsx')
