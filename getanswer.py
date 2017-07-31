import xlrd
def getstart():
    book = xlrd.open_workbook("myexcel.xls")
    name = input("Looking for which test?(all lower case, example: 73g)\nPress enter when you are done\ninput: ")
    for i in range(book.nsheets):
        sh = book.sheet_by_index(i)
        if name.upper() in sh.name or name.lower() in sh.name:
            print (name, i)
            p = i
            break
    return ([name.lower(),p])

def getanswer(p):
    book = xlrd.open_workbook("myexcel.xls")
    sh = book.sheet_by_index(p)
    
    answers = []
    for i in range(40):
        answers.append(sh.cell_value(rowx=i+2,colx=1))
    for i in range(35):
        answers.append(sh.cell_value(rowx=i+2,colx=3))

    for i in range(40):
        answers.append(sh.cell_value(rowx=i+2,colx=5))
    for i in range(20):
        answers.append(sh.cell_value(rowx=i+2,colx=7))

    for i in range(40):
        answers.append(sh.cell_value(rowx=i+2,colx=9))

    for i in range(40):
        answers.append(sh.cell_value(rowx=i+2,colx=11))

    return (answers)

def main():
    book = xlrd.open_workbook("myexcel.xls")
    print ("The number of tests in answer sheet is", book.nsheets)
    p = 0
    result = getstart()
    name = result[0]
    p = result[1]
    answer=getanswer(p)
    print ("Below is the answer set for", name)
    print (answer)
    return (answer)


main()
again = input("wanna run another time?(y/n)")
if (again=="y"):
    main()
