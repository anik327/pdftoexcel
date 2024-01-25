import os
import PyPDF2
import openpyxl
import re


pdfFileList = []
FileList = os.listdir(os.getcwd())
pdfRegx = re.compile(r'[^\\]*\.pdf')
for file in FileList:
    if (pdfRegx.search(file)):
        pdfFileList.append(file)

data = []

for pdf in pdfFileList:
    pdfFileObj = open(pdf, 'rb')
    pdfReader = PyPDF2.PdfReader(pdfFileObj)
    pageNumber = len(pdfReader.pages)
    pageObj = pdfReader.pages[0]
    extractedData = pageObj.extract_text()

    jobDocketNoExtracRegex = re.compile(r'(JC\/\d\d-\d\d\/\d\d\d\d)')
    jobDockNo = jobDocketNoExtracRegex.search(extractedData).group()

    jobDocketDateExtracRegex = re.compile(r'JOB DOCKET DATE (\d{2}\/\d{2}\/\d{4})')
    jobDocketDate = jobDocketDateExtracRegex.search(extractedData)
    data.append([jobDockNo, jobDocketDate.group(1)])

wb = openpyxl.Workbook()  #Creates a blank workbook
sheet = wb.active
sheet.append(["JOB DOCKET NO.", "JOB DOCKET DATE", "JOB NAME", "ORD.QTY", "PRINT QTY", "PRINT QTY WITH WASTAGE", "DIMENSIONS", "Prod. Type", "Collar Flap", "Side Paste", "Tuck Flap", "Order Qty", "Extra Qty", "Print Qty", "Wastage Qty", "Final Job Qty", "Board Details", "UPS", "SHEETS", "KG", "CUT SIZE", "CUT SHEET", "WASTAGE SHEET", "CUTS", "CUT SHEET UPS", "TOTAL CUT SHEETS", "WASTAGE"])
for row in data:
    sheet.append(row)
wb.save('dataFile.xlsx')
