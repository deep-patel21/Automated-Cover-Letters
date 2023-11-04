import os
import sys
import docx
import shutil
from pathlib import Path
from docx.shared import Pt
from datetime import date
from docx2pdf import convert
from openpyxl import load_workbook as loadwb

class ExcelRow:
    def __init__(self, company, position, requisitionID, contact):
        self.company = company
        self.position = position
        self.requisitionID = requisitionID
        self.contact = contact

    def __str__(self):
        return str(self.company) + " " + str(self.position) + " " + str(self.requisitionID) + " " + str(self.contact)

def manageDocs(company, position):
    pathOfCompany = company.replace(" ", "")
    pathString = os.path.join(os.getcwd(), pathOfCompany)
    Path(pathString).mkdir(parents=True, exist_ok=True)
    pathOfPosition = position.replace(" ", "")
    pathOfTemplate = os.path.join(os.getcwd(), 'cl_template.docx')
    pathOfDestination = os.path.join(pathString, pathOfPosition + '.docx')
    shutil.copy(pathOfTemplate, pathOfDestination)
    return pathOfDestination

def replace_string(filename, key, replacement):
    document = docx.Document(filename)
    style = document.styles['Normal']
    font = style.font
    font.name = 'Times\ New\ Roman'
    font.size = Pt(12)

    for paragraph in document.paragraphs:
        if key in paragraph.text:
            print('Replacement Located.')
            text = paragraph.text.replace(key, replacement)
            paragraph.text = text
            paragraph.style = document.style['Normal']
    document.save(filename)

def convertToPDF(destination):
    convert(destination)

#Read data from Excel File containing TK
def readXL():
    ExcelRows = []
    myFile = loadwb(filename='cl_data.xlsx')
    sh = myFile.active

    for fileRows in range(1, sh.max_row + 1):
        row = []
        for fileCols in range(1, sh.max_column + 1):
            cell = sh.cell(row = fileRows, column = fileCols)
            row.append(str(cell.value))
        content = ExcelRow(row[0], row[1], row[2], row[3])
        ExcelRows.append(content)
    return ExcelRows

#Create one PDFs for each row read from readXL()
def batchGenerate(letter):
    for row in letter:
        company = row.company
        position = row.position
        requsitionID = row.requisitionID
        contact = row.contact

        if requsitionID == 'default':
            requsitionID = ""
        else:
            requsitionID = " of Req.ID " + requsitionID
        if contact == 'default':
            contact == "Hiring Manager"

        dict = {'#companyName#': company,
                    '#date#': date.today().strftime("%B %d, %Y"),
                    '#jobTitle#': position,
                    '#jobId#': requsitionID,
                    '#contactName#':contact}

        destination = manageDocs(company, position)

        for i in dict:
            replace_string(destination, i, dict[i])

        convertToPDF(destination)

if __name__ == "__main__":
  os.chdir(os.path.dirname(sys.argv[0]))
  letter = readXL()
  batchGenerate(letter)