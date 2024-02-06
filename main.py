from src.xlsx_editor import GetArray
# import src.docx_editor
from docx2pdf import convert

file_name = "files/source.xlsx"
sheet_name = "Samuel y Pablo"

print(GetArray(file_name, sheet_name))


# file_name = "files/test"
# convert(file_name + ".docx")
# convert(file_name + ".docx", file_name + ".pdf")
# convert("files/")