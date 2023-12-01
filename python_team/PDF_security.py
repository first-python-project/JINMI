import PyPDF2
import os

RPath = "D:\\python_team\\Result\\워드 파일1.pdf"

pdf = PyPDF2.PdfReader(open(RPath, "rb"))

password = input("암호를 입력하세요: ")

encrypted_pdf = PyPDF2.PdfWriter()
encrypted_pdf.append_pages_from_reader(pdf)
encrypted_pdf.encrypt(password)

SPath = "D:\\python_team\\Secret\\encrypted.pdf"

with open(SPath, "wb") as output_file:
    encrypted_pdf.write(output_file)


