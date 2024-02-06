from openpyxl import load_workbook

def GetArray(file_name, sheet_name):

    array = []

    workbook = load_workbook(file_name)

    sheet = workbook.active

    for i in range(4,213):
        if (sheet[f"A{i}"].value == 2 and i == 13):
            dictionary = {}
            dictionary["id"] = sheet[f"B{i}"].value
            dictionary["nombre"] = sheet[f"D{i}"].value
            dictionary["instituto"] = sheet[f"H{i}"].value
            dictionary["provincia"] = sheet[f"I{i}"].value
            dictionary["eso"] = sheet[f"J{i}"].value
            dictionary["bachillerato"] = sheet[f"K{i}"].value
            dictionary["ingles"] = sheet[f"L{i}"].value
            dictionary["extraescolares"] = sheet[f"M{i}"].value
            dictionary["ensayo1"] = sheet[f"N{i}"].value
            dictionary["ensayo2"] = sheet[f"O{i}"].value
            dictionary["grado"] = sheet[f"Q{i}"].value
            dictionary["comentarios"] = sheet[f"R{i}"].value
            array.append(dictionary)

    return array