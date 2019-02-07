from openpyxl import load_workbook

#loading excel wb and declaring the sheet
def loadWorkbook(filename):
    wb = load_workbook(filename) #= 'DATA_BUKU.xlsx')
    sheet1 = wb['Sheet1']

#prompting for cell value
def promptNo_siri():
    no_siri = input("No siri : ")

#finding cell and its row
def findCell(sheet1, no_siri):
    i = 1
    for row in sheet1.iter_rows():
        if sheet1.cell(row = i,column = 1).value == int(no_siri) :
            print(sheet1.cell(row = i,column = 1))
            #printing the excel first row (header)
            list_with_values_header=[]
            for cell in sheet1[1]:
                list_with_values_header.append(cell.value)
            print(list_with_values_header)
            #printing the row
            list_with_values=[]
            for cell in sheet1[i]:
                list_with_values.append(cell.value)
            print(list_with_values)
            print(sheet1.cell(row = i,column = 1).coordinate)
            break
        i += 1

#checking for correct value.
def cellCheck(sheet1, i):
    cell = sheet1.cell(row = i,column = 1).coordinate
    print("value for cell " + cell + ": " + str(sheet1[cell].value)) 

#list used to turn string to char array
#print(list(cell))

#changing to tick coord
def tickCoord(sheet1, cell):
    # row and col char are separated
    char1 = list(cell)[0]
    char2 = list(cell)[1]
    #checking and incrementing it to be at 'Tick" column
    print(char1)
    charnew = chr(ord(char1) + 5)
    print(charnew)
    #joining existing constant row char with new col value
    inputnew = ''.join(charnew + char2)
    print(inputnew)
    #checking existing value at Tick column
    print(sheet1[inputnew].value) 

def insertTick(sheet1, inputnew):
    #inserting a 'tick' ("/")
    sheet1[inputnew] = "/"
    #checking its value now
    print(sheet1[inputnew].value) 

#saving the file
def saveWorkbook(wb, filename):
    #make sure to close the excel or else perm is denied
    wb.save(filename) #= 'DATA_BUKU.xlsx')