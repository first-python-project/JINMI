import os
import re
import win32com.client
from docx2pdf import convert
from fpdf import FPDF
import PyPDF2

TDPath = "D:\\python_team\\ToDo" 
RPath = "D:\\python_team\\Result"

#PPT to PDF 만들어주는 함수
def ppt2pdf():
    files = [f for f in os.listdir(TDPath) if re.match('.*[.]pptx', f)]
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    for file in files:
        # PPT 파일을 PDF로 바꾸는 로직
        deck = powerpoint.Presentations.Open(os.path.join(TDPath, file))
        pre, ext = os.path.splitext(file)
        deck.SaveAs(os.path.join(RPath, pre + ".pdf"), 32)  # formatType = 32 for ppt to pdf
        deck.Close()

    powerpoint.Quit()

ppt2pdf()

#Word to PDF 만들어주는 함수
def word2pdf():
    files = [f for f in os.listdir(TDPath) if re.match('.*[.]docx', f)]
    for file in files:
        # Word 파일을 PDF로 바꾸는 로직
        pre, ext = os.path.splitext(file)
        convert(os.path.join(TDPath, file), os.path.join(RPath, pre + ".pdf"))

word2pdf()

#HWP to PDF 만들어주는 함수
def hwp2pdf(): 
    files = [f for f in os.listdir(TDPath) if re.match('.*[.]hwp', f)]
    hwp = win32com.client.gencache.EnsureDispatch('HWPFrame.HwpObject')
    hwp.RegisterModule('FilePathCheckDLL', 'SecurityModule')
    for file in files:
        # HWP 파일을 PDF로 바꾸는 로직
        hwp.Open(os.path.join(TDPath, file))
        pre, ext = os.path.splitext(file)
        hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = os.path.join(RPath, pre + ".pdf")
        hwp.HParameterSet.HFileOpenSave.Format = "PDF"
        hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet);

    hwp.Quit()

hwp2pdf()

# PDF 파일 암호 생성 하기
result_path = 'Result'
secret_path = 'Secret'

def encrypt_pdf():
    for file in os.listdir(result_path):
        if file.endswith(".pdf"):
            input_pdf = os.path.join(result_path, file)

            # Enter PDF file password
            pdf = PyPDF2.PdfReader(open(input_pdf, "rb"))
            password = input(f"Enter the password for '{file}': ")

            # Encrypt the PDF
            encrypted_pdf = PyPDF2.PdfWriter()
            encrypted_pdf.append_pages_from_reader(pdf)
            encrypted_pdf.encrypt(password)

            # Save the encrypted PDF in the "Secret" folder with the file name + "_encrypted"
            encrypted_file = os.path.splitext(file)[0] + "_encrypted.pdf"
            output_pdf = os.path.join(secret_path, encrypted_file)

            with open(output_pdf, "wb") as output_file:
                encrypted_pdf.write(output_file)

            print(f"Encrypted file saved: {secret_path}")


encrypt_pdf()