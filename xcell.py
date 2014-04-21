import xlrd
import xlwt

path = "C:\Users\Yashovardhan\Downloads\Proj6-ModellingHumanIntuition.xlsx"
path2 = "C:\Users\Yashovardhan\Downloads\Proj6-ModellingHumanIntuition(Team 7).xls"
row_nums = [5,44,51,70,80,84,133,137,151,160,178,190,196,203,207,237,239,321,328,358,368,381,413,421,439,449,494,497,506,514,533,537,562,563,579,598,605,612,619,677,679,690,702,752,776,779,819,855,862,887,903,904,979,989,1011,1031,1044,1143,1159,1163,1165,1166,1196,1198,1218,1219,1243,1291,1294,1330,1342,1358,1361,1363,1414,1430,1436,1447,1469,1494,1528,1533,1539,1555,1595,1598,1620]
book = xlrd.open_workbook(path)
wb = xlwt.Workbook()

def PageOne():
    for i in range(1):
        sh = book.sheet_by_index(i)
        ws = wb.add_sheet('Sheet1')

        ws.write(0,0,'Name')
        ws.write(0,1,'Email')
        ws.write(0,2,'Gender')
        ws.write(0,3,'Age')
        ws.write(0,4,'Qualifications')
        ws.write(0,5,'Team ID')
        k = 1

        for j in row_nums:
            #print "%s\t\t %s\t\t %s\t %s\t %s\t\t %s\t\n" % (sh.cell_value(rowx=j-1,colx=0), sh.cell_value(rowx=j-1,colx=1), sh.cell_value(rowx=j-1,colx=2), sh.cell_value(rowx=j-1,colx=3), sh.cell_value(rowx=j-1,colx=4), sh.cell_value(rowx=j-1,colx=5))
            ws.write(k,0,sh.cell_value(rowx=j-1,colx=0))
            ws.write(k,1,sh.cell_value(rowx=j-1,colx=1))
            ws.write(k,2,sh.cell_value(rowx=j-1,colx=2))
            ws.write(k,3,sh.cell_value(rowx=j-1,colx=3))
            ws.write(k,4,sh.cell_value(rowx=j-1,colx=4))
            ws.write(k,5,sh.cell_value(rowx=j-1,colx=5))
            k = k + 1
    wb.save(path2)

def PageTwo():
    sh = book.sheet_by_index(1)
    ws = wb.add_sheet('Sheet2')
    col_num = 0

    for i in range(10):
        ws.write(0,col_num,'Question Asked')
        col_num+=1
        ws.write(0,col_num,'Image Given')
        col_num+=1
        ws.write(0,col_num,'User Answer')
        col_num+=1
        ws.write(0,col_num,'User Text Answer')
        col_num = col_num + 2
    ws.write(0,col_num,'Time taken')

    k = 1

    for j in row_nums:
        col_num = 0
        for x in range(10):
            ws.write(k,col_num,sh.cell_value(rowx=j-1,colx=col_num))
            col_num+=1
            ws.write(k,col_num,sh.cell_value(rowx=j-1,colx=col_num))
            col_num+=1
            ws.write(k,col_num,sh.cell_value(rowx=j-1,colx=col_num))
            col_num+=1
            ws.write(k,col_num,sh.cell_value(rowx=j-1,colx=col_num))
            col_num = col_num + 2
        ws.write(k,col_num,sh.cell_value(rowx=j-1,colx=col_num))
        k = k + 1
    wb.save(path2)

def PageThree():
    sh = book.sheet_by_index(2)
    ws = wb.add_sheet('Sheet3')
    col_num = 0

    for i in range(10):
        ws.write(0,col_num,'Question Asked')
        col_num+=1
        ws.write(0,col_num,'Image Given')
        col_num+=1
        ws.write(0,col_num,'User Answer')
        col_num+=1
        ws.write(0,col_num,'User Text Answer')
        col_num = col_num + 2
    ws.write(0,col_num,'Time taken')

    k = 1

    for j in row_nums:
        col_num = 0
        for x in range(10):
            ws.write(k,col_num,sh.cell_value(rowx=j-1,colx=col_num))
            col_num+=1
            ws.write(k,col_num,sh.cell_value(rowx=j-1,colx=col_num))
            col_num+=1
            ws.write(k,col_num,sh.cell_value(rowx=j-1,colx=col_num))
            col_num+=1
            ws.write(k,col_num,sh.cell_value(rowx=j-1,colx=col_num))
            col_num = col_num + 2
        ws.write(k,col_num,sh.cell_value(rowx=j-1,colx=col_num))
        k = k + 1
    wb.save(path2)

def PageFour():
    sh = book.sheet_by_index(3)
    ws = wb.add_sheet('Sheet4')
    col_num = 0

    for i in range(10):
        ws.write(0,col_num,'Question')
        col_num+=1
        ws.write(0,col_num,'User Answer')
        col_num+=1

    ws.write(0,col_num,'Time taken')

    k = 1

    for j in row_nums:
        col_num = 0
        for x in range(10):
            ws.write(k,col_num,sh.cell_value(rowx=j-1,colx=col_num))
            col_num+=1
            ws.write(k,col_num,sh.cell_value(rowx=j-1,colx=col_num))
            col_num+=1

        ws.write(k,col_num,sh.cell_value(rowx=j-1,colx=col_num))
        k = k + 1
    wb.save(path2)

def PageFive():
    sh = book.sheet_by_index(4)
    ws = wb.add_sheet('Sheet5')
    col_num = 0

    for i in range(10):
        ws.write(0,col_num,'Question')
        col_num+=1
        ws.write(0,col_num,'User Answer')
        col_num+=1

    ws.write(0,col_num,'Time taken')

    k = 1

    for j in row_nums:
        col_num = 0
        for x in range(10):
            ws.write(k,col_num,sh.cell_value(rowx=j-1,colx=col_num))
            col_num+=1
            ws.write(k,col_num,sh.cell_value(rowx=j-1,colx=col_num))
            col_num+=1

        ws.write(k,col_num,sh.cell_value(rowx=j-1,colx=col_num))
        k = k + 1
    wb.save(path2)

def PageSix():
    sh = book.sheet_by_index(5)
    ws = wb.add_sheet('Sheet6')
    col_num = 0

    ws.write(0,col_num,'Question')
    col_num+=1

    for i in range(30):
        ws.write(0,col_num,'Answer')
        col_num+=1
        ws.write(0,col_num,'Time taken')
        col_num+=1

    k = 1

    for j in row_nums:
        col_num = 0
        ws.write(k,col_num,sh.cell_value(rowx=j-1,colx=col_num))
        col_num+=1

        for x in range(30):
            ws.write(k,col_num,sh.cell_value(rowx=j-1,colx=col_num))
            col_num+=1
            ws.write(k,col_num,sh.cell_value(rowx=j-1,colx=col_num))
            col_num+=1
        k = k + 1
    wb.save(path2)

def PageSeven():
    sh = book.sheet_by_index(6)
    ws = wb.add_sheet('Sheet7')
    col_num = 0

    ws.write(0,col_num,'Question')
    col_num+=1

    for i in range(30):
        ws.write(0,col_num,'Answer')
        col_num+=1
        ws.write(0,col_num,'Time taken')
        col_num+=1

    k = 1

    for j in row_nums:
        col_num = 0
        ws.write(k,col_num,sh.cell_value(rowx=j-1,colx=col_num))
        col_num+=1

        for x in range(30):
            ws.write(k,col_num,sh.cell_value(rowx=j-1,colx=col_num))
            col_num+=1
            ws.write(k,col_num,sh.cell_value(rowx=j-1,colx=col_num))
            col_num+=1
        k = k + 1
    wb.save(path2)

PageOne()
PageTwo()
PageThree()
PageFour()
PageFive()
PageSix()
PageSeven()
