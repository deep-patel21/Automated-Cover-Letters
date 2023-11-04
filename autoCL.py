import os
import sys
import docx
from pathlib import Path
from docx.shared import Pt
from datetime import date
from docx2pdf import convert
from openpyxl import load_workbook as loadwb

def readXL():
    rows = []
    myFile = loadwb(filename='cl.xlsx')
    sh = myFile.active

    for fileRows in range(1, sh.max_row + 1):
        row = []
        for fileCols in range(1, sh.max_column + 1):
            cell = sh.cell(row = fileRows, column = fileCols)
            row.append(str(cell.value))
        content = rows(row[0], row[1], row[2], row[3])
        rows.append(content)
    return rows

def batchGenerate():
    for row in letter:
        company = row.company
        position = row.position
        requsitionID = row.requisitionID
        contact = row.contact

if __name__ == "__main__":
  os.chdir(os.path.dirname(sys.argv[0]))
  letter = readXL()
  batchGenerate(letter)
