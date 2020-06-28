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

def getSum(rowsNum, tabulka):
    sumOfPrices = 0
    for i in range(1, rowsNum):
        sumOfPrices += float(tabulka["C" + str(i)].value)
    return sumOfPrices

def writeToLastRow(tabulka, lastrowindex, datum, jmeno, cena, workbook):
    tabulka["A" + str(lastrowindex)].value = datum
    tabulka["B" + str(lastrowindex)].value = jmeno
    tabulka["C" + str(lastrowindex)].value = cena 
    workbook.save("test.xlsx")
    dalsi = input("Pro pridani produktu zadej [y] a pro zobrazeni souctu cen napis [s]")
    if dalsi == "y":
        writeToLastRow(sheet, findLastRow(sheet),input("Zadej datum: "),input("\nZadej jmeno produktu: "),input("\nZadej cenu produktu: "), wb )
        print("Polozka uspesne pridana.")
    elif dalsi == "s":
        print(f"Součet všech položek je: {getSum((findLastRow(sheet)),sheet)} kc.")
        nevimjakejmenolul = input("Pro pridani produktu zadej [y]")
        if nevimjakejmenolul == "y":
            writeToLastRow(sheet, findLastRow(sheet),input("Zadej datum: "),input("\nZadej jmeno produktu: "),input("\nZadej cenu produktu: "), wb )
            print("Polozka uspesne pridana.")
        else:
            pass
    else:
        pass


#print(getSum(findLastRow(sheet), sheet))
#writeToLastRow(sheet, findLastRow(sheet),input("Zadej datum: "),input("\nZadej jmeno produktu: "),input("\nZadej cenu produktu: "), wb )
base = input("Pro pridani produktu zadej [y] a pro zobrazeni souctu cen napis [s]")
if base == "y":
    print("Polozka uspesne pridana.")
elif base == "s":
    print(f"Součet všech položek je: {getSum((findLastRow(sheet)),sheet)}.")
    nevimjakejmenolul2 = input("Pro pridani produktu zadej [y] ")
    if nevimjakejmenolul2 == "y":
        writeToLastRow(sheet, findLastRow(sheet),input("Zadej datum: "),input("\nZadej jmeno produktu: "),input("\nZadej cenu produktu: "), wb )
        print("Polozka uspesne pridana.")
    else:
        pass
else:
    pass



