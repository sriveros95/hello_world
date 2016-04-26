import xlrd
import time
import sys

## hankDaBot.py "decimalList.xls" 0 dec-hex
##                  xls file    column  conversion

file = sys.argv[1]
column = int(sys.argv[2])
op = sys.argv[3]


if sys.argv[1] == "dec-hex":
    print("Convertire Decimales a Hexadecimal")
    
def dec_to_hex(decimalNum):
    try:
        prueba = decimalNum + 1
    except ValueError:
        print("Not a Decimal!!!")
        return -1
    hexa = []
    stop = 0
    n = 0
    while stop == 0:
        if(int(decimalNum)>=(16**n)):
            n = n + 1
        else:
            stop = 1
            n = n - 1
    actual = decimalNum
    while n >= 0:
        hex_element = actual//(16**n)
        if((hex_element >= 0) and (hex_element <= 9)):
            hexa.append(actual//(16**n))            
        elif hex_element == 10:
            hexa.append('A')
        elif hex_element == 11:
            hexa.append('B')
        elif hex_element == 12:
            hexa.append('C')
        elif hex_element == 13:
            hexa.append('D')
        elif hex_element == 14:
            hexa.append('E')
        elif hex_element == 15:
            hexa.append('F')
        else:
            return -1
        actual = actual % (16**n)
        n = n - 1
        hex_string = ''.join(str(v) for v in hexa)

    print(" hexadecimal es : ", hex_string)

def dec_to_oct(decimalNum):
    try:
        prueba = decimalNum + 1
    except ValueError:
        print("Not a Decimal!!!")
        return -1
    octal = []
    stop = 0
    n = 0
    while stop == 0:
        if(int(decimalNum)>=(8**n)):
            n = n + 1
        else:
            stop = 1
            n = n - 1
    actual = decimalNum
    while n >= 0:
        octal_element = actual//(8**n)
        if((octal_element >= 0) and (octal_element <= 8)):
            octal.append(actual//(8**n))            
        else:
            return -1
        actual = actual % (8**n)
        n = n - 1
        octal_string = ''.join(str(v) for v in octal)

    print(" octal es : ", octal_string)


workbook = xlrd.open_workbook(file)
worksheet = workbook.sheet_by_index(0)
i = 0
stop = 0
if op == "dec-hex":
    while(stop == 0):
        try:
            current_cell = int(worksheet.cell(i, column).value)
            if current_cell == xlrd.empty_cell.value:
                stop = 1
            else:
                print(current_cell, end = "")
                dec_to_hex(current_cell)
                i = i + 1
                
        except IndexError:
            stop = 1
        except ValueError:
            print("!  ", worksheet.cell(i, column).value, end='')
            print(" is not a Decimal!!!")
            i = i + 1
        
    print("Bye bye!")

elif op == "dec-oct":
    while(stop == 0):
        try:
            current_cell = int(worksheet.cell(i, column).value)
            if current_cell == xlrd.empty_cell.value:
                stop = 1
            else:
                print(current_cell, end = "")
                dec_to_oct(current_cell)
                i = i + 1
                
        except IndexError:
            stop = 1
        except ValueError:
            print("!  ", worksheet.cell(i, column).value, end='')
            print(" is not a Decimal!!!")
            i = i + 1
        
    print("Bye bye!")
    pass
