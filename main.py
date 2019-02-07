from openpyxl import load_workbook
from prettytable import PrettyTable

#loading excel wb and declaring the sheet
excfile = 'DATA_BUKU.xlsx'
wb = load_workbook(filename = excfile)
sheet1 = wb['Sheet1']


#prompting for action
def promptAction():
    global action
    print("KEY \n1.INSERT TICK\n2.DELETE TICK\n3.VIEW ROW\n4.ANALYSIS\n5.QUICK TICK\n6.QUIT")
    action = input("Enter a key :")
    
    if action == '1' :
        print('You chose to insert a tick to a book row.')
        act12()
    elif action == '2' :
        print('You chose to delete a tick from a book row.')
        act12()
    elif action == '3' :
        print('You chose to view the row')
        act3()
    elif action == '4' :
        print("I'll give you an overview analysis")
    elif action == '5' :
        print("You chose quick tick, be careful... C: \nEnter 'Q' to quit.")
        act5()
    elif action == '6' :
        wb.save(filename = excfile)
        exit()
    else :
        print('Your input is out of context, please try again.')
        promptAction()

#action 1 and 2
def act12():
    promptNo_siri()
    findCell(sheet1, no_siri)
    cellCheck(sheet1, i)
    tickCoord(sheet1, cell)
    if action == '1' :
        insertTick(sheet1, inputnew)
        newTable = PrettyTable()
        newTable.field_names = (list_with_values_header)
        list_with_values[5] = "/"
        newTable.add_row(list_with_values)
        print(newTable)
        wb.save(filename = excfile)
    elif action == '2' :
        deleteTick(sheet1, inputnew)
        newTable = PrettyTable()
        newTable.field_names = (list_with_values_header)
        list_with_values[5] = ""
        newTable.add_row(list_with_values)
        print(newTable)
        wb.save(filename = excfile)
    else :
        wb.save(filename = excfile)
        exit()
        
def act3():
    promptNo_siri()
    global i
    i = 1
    for row in sheet1.iter_rows():
        if sheet1.cell(row = i,column = 1).value == int(no_siri) :
            #   print(sheet1.cell(row = i,column = 1))
            #printing the excel first row (header)
            list_with_values_header=[]
            for cell in sheet1[1]:
                list_with_values_header.append(cell.value)
            #print(list_with_values_header)
            #printing the row
            list_with_values=[]
            for cell in sheet1[i]:
                list_with_values.append(cell.value)
            #print(list_with_values)
            #   print(sheet1.cell(row = i,column = 1).coordinate)
            break
        i += 1
    newTable = PrettyTable()
    newTable.field_names = (list_with_values_header)
    newTable.add_row(list_with_values)
    print(newTable)

def act5():
    while True :
            global i
            promptNo_siri()
            if no_siri == 'Q':
                break
            findCell(sheet1, no_siri)
            cellCheck(sheet1, i)
            if i == 420 :
                print("No book found, please check for typos or use another number.")
                promptNo_siri()
            elif i > 420 : 
                print("No book found, please check for typos or use another number.")
                promptNo_siri()
            else:
                tickCoord(sheet1, cell)
                insertTick(sheet1, inputnew)
                wb.save(filename = excfile)
                newTable = PrettyTable()
                newTable.field_names = (list_with_values_header)
                list_with_values[5] = "/"
                newTable.add_row(list_with_values)
                print(newTable)

#prompting for cell value
def promptNo_siri():
    global no_siri
    no_siri = input("No siri : ")
    try:
        no_siri = int(no_siri)
        if no_siri < 0 :
            print("Only positive number please.")
            promptNo_siri()
        else :
            True
    except ValueError:
        if no_siri == 'Q':
            promptAction()
        else :
            print("That's not an int!")
            promptNo_siri() 

#finding cell and its row
def findCell(sheet1, no_siri):
    global i, list_with_values_header, list_with_values
    i = 1
    for row in sheet1.iter_rows():
        if sheet1.cell(row = i,column = 1).value == int(no_siri) :
            #   print(sheet1.cell(row = i,column = 1))
            #printing the excel first row (header)
            list_with_values_header=[]
            for cell in sheet1[1]:
                list_with_values_header.append(cell.value)
            #   print(list_with_values_header)
            #printing the row
            list_with_values=[]
            for cell in sheet1[i]:
                list_with_values.append(cell.value)
            #   print(list_with_values)
            #   print(sheet1.cell(row = i,column = 1).coordinate)
            
            break
        i += 1
    if i == 420 :
        print("No book found, please check for typos or use another number.")
        promptNo_siri()

#checking for correct value.
def cellCheck(sheet1, i):
    global cell
    cell = sheet1.cell(row = i,column = 1).coordinate
    print("value for cell " + cell + ": " + str(sheet1[cell].value)) 

#list used to turn string to char array
#print(list(cell))

#changing to tick coord
def tickCoord(sheet1, cell):
    # row and col char are separated
    char1 = list(cell)[0]
    char2 = list(cell)[1]
    if bool(list(cell)[2]) == True:
        char3 = list(cell)[2]
    elif bool(list(cell)[3]) == True:
        char4 = list(cell)[3]
    else:
        True
    #checking and incrementing it to be at 'Tick" column
    #   print(char1)
    charnew = chr(ord(char1) + 5)
    #   print(charnew)
    #joining existing constant row char with new col value
    global inputnew
    inputnew = ''.join(charnew + char2)
    if bool(list(cell)[2]) == True:
        inputnew = ''.join(charnew + char2 + char3)
    elif bool(list(cell)[3]) == True:
        inputnew = ''.join(charnew + char2 + char3 + char4)
    else:
        True
    #   print(inputnew)
    #checking existing value at Tick column
    print("Coord " + inputnew +" before : "+ str(sheet1[inputnew].value) )

def insertTick(sheet1, inputnew):
    #inserting a 'tick' ("/")
    sheet1[inputnew] = "/"
    #checking its value now
    print("Coord " + inputnew +" after : "+ str(sheet1[inputnew].value)  )

def deleteTick(sheet1, inputnew):
    #deleting a 'tick' ("/")
    sheet1[inputnew] = ""
    #checking its value now
    print("Coord " + inputnew +" after : "+ str(sheet1[inputnew].value)  )

while True :
    promptAction()


#saving the file
#make sure to close the excel or else perm is denied
wb.save(filename = excfile)
