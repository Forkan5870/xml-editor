from src.xlsx_editor import GetArray
from src.docx_editor import CreateDocx
from docx2pdf import convert

file_name = "files/source.xlsx"
numero_tutor = 2

array = GetArray(file_name, numero_tutor)
CreateDocx(array)

file_name = "files/test"
convert(file_name + ".docx")
convert(file_name + ".docx", file_name + ".pdf")
convert("files/")