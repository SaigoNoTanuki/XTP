import os
import json
import xlwings as xw

jsonEmpty = False

paths = {}

def PopulateJSON():
    setWbPath = input("Enter path to Work Books. \n")
    setPdfPath = input("Enter path to save PDFs. \n")

    setConfig = {
    "wbPath": setWbPath,
    "pdfPath": setPdfPath
    }

    with open("config.json", 'w') as path:
        json.dump(setConfig, path)

with open("config.json", "r") as path:
    try:
        paths = json.load(path)
    except Exception as err:
        print(f"JSON empty: {err}")
        jsonEmpty = True

if(jsonEmpty == True):
    PopulateJSON()
    jsonEmpty = False
    with open("config.json", "r") as path:
        paths = json.load(path)
else:
    cmnd = input("Enter to continue \n (1) to change paths \n")
    if(cmnd == "1"):
        PopulateJSON()
        with open("config.json", "r") as path:
            paths = json.load(path)

app = xw.App(visible=False)

if(jsonEmpty == False):
    for file in os.scandir(paths["wbPath"]):
        if file.name.endswith(".xlsx"):
            try:
                wb = app.books.open(file.path)
                wb.api.ExportAsFixedFormat(0, f"{paths['pdfPath']}/{os.path.splitext(file.name)[0]}.pdf")
                wb.close()
            except Exception as err:
                print(err)

app.quit()