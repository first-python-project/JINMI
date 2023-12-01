import os
import re
import win32com.client
from docx2pdf import convert
from fpdf import FPDF
import PyPDF2

powerpoint = win32com.client.Dispatch("Powerpoint.Application")
hwp = win32com.client.gencache.EnsureDispatch('HWPFrame.HwpObject')
hwp.RegisterModule('FilePathCheckDLL', 'SecurityModule')
pdf = FPDF('L')

TDPath = "D:\\python_team\\ToDo" 
RPath = "D:\\python_team\\Result"

#PPT to PDF
files = [f for f in os.listdir(TDPath) if re.match('.*[.]pptx', f)]
for file in files:
    # PPT 파일을 PDF로 바꾸는 로직
    deck = powerpoint.Presentations.Open(os.path.join(TDPath, file))
    pre, ext = os.path.splitext(file)
    deck.SaveAs(os.path.join(RPath, pre + ".pdf"), 32)  # formatType = 32 for ppt to pdf
    deck.Close()

powerpoint.Quit()

#Word to PDF
files = [f for f in os.listdir(TDPath) if re.match('.*[.]docx', f)]
for file in files:
    # Word 파일을 PDF로 바꾸는 로직
    pre, ext = os.path.splitext(file)
    convert(os.path.join(TDPath, file), os.path.join(RPath, pre + ".pdf"))


#HWP to PDF
files = [f for f in os.listdir(TDPath) if re.match('.*[.]hwp', f)]
for file in files:
    # HWP 파일을 PDF로 바꾸는 로직
    hwp.Open(os.path.join(TDPath, file))
    pre, ext = os.path.splitext(file)
    hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
    hwp.HParameterSet.HFileOpenSave.filename = os.path.join(RPath, pre + ".pdf")
    hwp.HParameterSet.HFileOpenSave.Format = "PDF"
    hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet);

hwp.Quit()

RSPath = "D:\\python_team\\Result\\.*[.]pdf"

pdf = PyPDF2.PdfReader(open(RSPath, "rb"))

password = input("암호를 입력하세요: ")

encrypted_pdf = PyPDF2.PdfWriter()
encrypted_pdf.append_pages_from_reader(pdf)
encrypted_pdf.encrypt(password)

SPath = 'D:\\python_team\\Secret\\encrypted.pdf'

with open(SPath, "wb") as output_file:
    encrypted_pdf.write(output_file)
