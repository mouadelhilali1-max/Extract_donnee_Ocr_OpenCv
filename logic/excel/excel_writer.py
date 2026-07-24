from openpyxl import Workbook


def save_to_excel(data, filename="SpecificationTree.xlsx"):

    wb = Workbook()

    ws = wb.active

    ws.title = "Tree"

    ws.append(["Texte", "X", "Y", "Confiance"])

    for ligne in data:

        ws.append([
            ligne["text"],
            ligne["x"],
            ligne["y"],
            ligne["confidence"]
        ])

    wb.save(filename)

    print("Excel sauvegardé :", filename)