import openpyxl
wb = openpyxl.load_workbook("test.xlsx")
sheet = wb.get_sheet_by_name("List1")
#print(wb.get_sheet_names())

print("-----------------")
print("PROGRAM STARTING")
print("-----------------")

def findLastRow(tabulka):
    main = True
    index = 1

    while main:
        if tabulka["A" + str(index)].value != None:
            index += 1
        else:
            main = False
            return index
def writeToLastRow(tabulka, lastrowindex, datum, jmeno, cena, workbook):
    tabulka["A" + str(lastrowindex)].value,tabulka["B" + str(lastrowindex)].value,tabulka["C" + str(lastrowindex)].value = datum, jmeno, cena + "kc"
    workbook.save("test.xlsx")
    dalsi = input("chcete zadat dalsi produkt? y/n")
    if dalsi == "y":
        writeToLastRow(sheet, findLastRow(sheet),input("Zadej datum: "),input("\nZadej jmeno produktu: "),input("\nZadej cenu produktu: "), wb )
        print("Polozka uspesne pridana.")
    else:
        pass
writeToLastRow(sheet, findLastRow(sheet),input("Zadej datum: "),input("\nZadej jmeno produktu: "),input("\nZadej cenu produktu: "), wb )







    
