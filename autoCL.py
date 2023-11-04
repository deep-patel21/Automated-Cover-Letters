import os
import sys
import docx
from pathlib import Path
from docx.shared import Pt
from datetime import date
from docx2pdf import convert
from openpyxl import load_workbook as loadwb

if __name__ == "__main__":
  os.chdir(os.path.dirname(sys.argv[0]))
  letter = readXL()
  batchGenerate(letterInfo)
