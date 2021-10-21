from openpyxl import load_workbook

def get_SPN(line):
    if ws.cell(line, 7).value is not None:
        return ws.cell(line, 7).value
    else:
        return 0

def get_FMI(line):
    if ws.cell(line, 8).value is not None:
        return ws.cell(line, 8).value
    else:
        return 255

def get_uiCodeNb(line):
    if ws.cell(line, 1).value is not None:
        return str(ws.cell(line, 1).value)
    else:
        return 0

def get_sPriority(line):
    if ws.cell(line, 3).value is not None:
        return str("'" + ws.cell(line, 3).value + "'")
    else:
        return 0

def get_usType(line):
    if ws.cell(line, 4).value is not None:
        return str(ws.cell(line, 4).value)
    else:
        return 0

def get_uiDeviceNb(line):
    if ws.cell(line, 5).value is not None:
        return str(ws.cell(line, 5).value)
    else:
        return str("1")

def get_sDescription(line):
    if ws.cell(line, 6).value is not None:
        return str("'" + ws.cell(line, 6).value + "'")
    else:
        return 0

def get_sDescriptionStaff():
    return str("''")

def get_sDescriptionMaintenance():
    return str("''")

try:
    wb = load_workbook('Bledy_silniki_Spalinowe_3.xlsx')
except:
    print("No such file")
    quit()

ws = wb.active
linesCount = 1
while(ws.cell(linesCount, 1).value is not None):
    linesCount += 1

try:
    codeFile = open('codeFile.txt', 'w')
except:
    print("File could not be created")
    quit()

print("(*SOFTCONTROL:\nVERSION:4.00.20*)\nFUNCTION_BLOCK DIAG_WyszukajKomunikatySilnika\n(**)\n(**)", file = codeFile)
print("\tVAR_INPUT", file = codeFile)
print("\t\tdi_SzukaneSPN: DINT:=0;", file = codeFile)
print("\t\tb_SzukaneFMI: BYTE:=0;", file = codeFile)
print("\tEND_VAR", file = codeFile)
print("\tVAR_OUTPUT", file = codeFile)
print("\t\tuiCodeNb: UINT:=0", file = codeFile)
print("\t\tsPriority: STRING[1]:=''", file = codeFile)
print("\t\tusType: USINT:=0", file = codeFile)
print("\t\tuiDeviceNb: UINT:=0", file = codeFile)
print("\t\tsDescription: STRING[300]:=''", file = codeFile)
print("\t\tsDescriptionStaff: STRING[1]:=''", file = codeFile)
print("\t\tsDescriptionMaintenance: STRING[1]:=''", file = codeFile)
print("\tEND_VAR", file = codeFile)
print("'ST'", file = codeFile)
print("BODY", file = codeFile)

caseBeginningStr = "CASE di_SzukaneSPN OF"
print(caseBeginningStr, file = codeFile)

previousSPN = 0

line = 2
while(line < linesCount):
    currentSPN = get_SPN(line)
    nextSPN = get_SPN(line + 1)

    caseLabelStr = str(currentSPN) + ":"
    print(caseLabelStr, file = codeFile)
    if nextSPN is not currentSPN:
        caseStr = "\tuiCodeNb := " + get_uiCodeNb(line) + ";\n" + "\tsPriority := " + get_sPriority(line) + ";\n" +  "\tusType := " + get_usType(line) + ";\n" + "\tuiDeviceNb := " + get_uiDeviceNb(line) + ";\n" + "\tsDescription := " + get_sDescription(line) + ";\n" + "\tsDescriptionStaff := " + get_sDescriptionStaff() + ";\n" + "\tsDescriptionMaintenance := " + get_sDescriptionMaintenance() + ";"
        print(caseStr, file = codeFile)
        line += 1
    else:
        caseBeginningStr = "\tCASE b_SzukaneFMI OF"
        print(caseBeginningStr, file = codeFile)

        newLine = line
        while((get_SPN(newLine) is get_SPN(newLine + 1)) or (get_SPN(newLine) is get_SPN(newLine - 1))):
            currentFMI = get_FMI(newLine)
            caseLabelStr = "\t" + str(currentFMI) + ":"
            print(caseLabelStr, file = codeFile)
            caseStr = "\t\tuiCodeNb := " + get_uiCodeNb(newLine) + ";\n" + "\t\tsPriority := " + get_sPriority(newLine) + ";\n" +  "\t\tusType := " + get_usType(newLine) + ";\n" + "\t\tuiDeviceNb := " + get_uiDeviceNb(newLine) + ";\n" + "\t\tsDescription := " + get_sDescription(newLine) + ";\n" + "\t\tsDescriptionStaff := " + get_sDescriptionStaff() + ";\n" + "\t\tsDescriptionMaintenance := " + get_sDescriptionMaintenance() + ";"
            print(caseStr, file = codeFile)
            print("\t\tBREAK;", file = codeFile)
            newLine += 1

        caseEndStr = "\tEND_CASE;"
        print(caseEndStr, file = codeFile)
        line = newLine
    caseBreak = "\tBREAK;"
    print(caseBreak, file = codeFile)
    previousSPN = get_SPN(line)

caseEndStr = "END_CASE;\nEND_BODY\nEND_FUNCTION_BLOCK"
print(caseEndStr, file = codeFile)