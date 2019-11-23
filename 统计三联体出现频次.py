import os
import openpyxl
from openpyxl.writer.excel import ExcelWriter

filePath = r'C:\Users\lenovo\Desktop'  # r'C:\Users\user\Desktop\找三联体bindingsite的规律\无二级结构的rna结构'
workBook = openpyxl.load_workbook(os.path.join(filePath, r'protein实验位点标记数据11.20.xlsx'), read_only=False)
sheetList = workBook.get_sheet_names()

expRes = {}
preRes = {}
# 颜色预设
red = openpyxl.styles.Font(name='Arial', size=11, bold=False, italic=False, vertAlign=None, underline='none',
                           strike=False, color='00FF0000')
black = openpyxl.styles.Font(name='Arial', size=11, bold=False, italic=False, vertAlign=None, underline='none',
                             strike=False, color='00000000')
green = openpyxl.styles.Font(name='Arial', size=11, bold=False, italic=False, vertAlign=None, underline='none',
                             strike=False, color='0000FF00')

columnC = 3  # expriment bindingsite
columnI = 9  # triplet
columnJ = 10  # prediction
columnM = 13  # pdb seq


def rnaNameAppend(expPattern, prePattern, dict, rnaName, tripletName):
    index = 0
    for j in range(prePattern):
        index += dict[tripletName][expPattern][j]
    dict[tripletName][expPattern].insert(index + 8, rnaName)


def printResultXls(row, rowIndex, column, pattern, triplet, color):
    ws.cell(row + rowIndex, column + 4).value = triplet[0]
    ws.cell(row + rowIndex, column + 5).value = triplet[1]
    ws.cell(row + rowIndex, column + 6).value = triplet[2]

    if pattern >> 2 & 1 == 1:
        ws.cell(row + rowIndex, column + 4).font = color
    else:
        pass

    if pattern >> 1 & 1 == 1:
        ws.cell(row + rowIndex, column + 5).font = color
    else:
        pass

    if pattern & 1 == 1:
        ws.cell(row + rowIndex, column + 6).font = color
    else:
        pass


for i in sheetList:
    print(i)

    curSheet = workBook[i]
    tripletPattern = 0
    curRow = 3
    preRow = 3
    # 向res 中添加rna名称 和 三联体
    while curRow < curSheet.max_row:
        if curSheet.cell(curRow, columnI).value == 1:
            if curRow - preRow == 3:
                tripletName = curSheet.cell(curRow - 2, columnM).value + curSheet.cell(curRow - 1, columnM).value + \
                              curSheet.cell(curRow, columnM).value
                # 统计实验模式
                tripletPattern = int(curSheet.cell(curRow - 2, columnC).value) * 4 + \
                                 int(curSheet.cell(curRow - 1, columnC).value) * 2 + \
                                 int(curSheet.cell(curRow, columnC).value)
                expRes.setdefault(tripletName, [[], [], [], [], [], [], [], []])[tripletPattern].append(i)
                # 预测位点统计
                prePattern = int(curSheet.cell(curRow - 2, columnJ).value) * 4 + \
                             int(curSheet.cell(curRow - 1, columnJ).value) * 2 + \
                             int(curSheet.cell(curRow, columnJ).value)

                preRes.setdefault(tripletName, [[0 for i in range(8)], [0 for i in range(8)], [0 for i in range(8)],
                                                [0 for i in range(8)], [0 for i in range(8)], [0 for i in range(8)],
                                                [0 for i in range(8)], [0 for i in range(8)]])[tripletPattern][
                    prePattern] += 1
                # preRes[tripletName][tripletPattern].append(i)
                rnaNameAppend(tripletPattern, prePattern, preRes, i, tripletName)
                preRow += 1
        else:
            preRow = curRow

        curRow += 1
        tripletPattern = 0

# 输出excel表格
# for i in preRes:
#    print(preRes)
resWorkBook = openpyxl.Workbook()
# ew=ExcelWriter(workbook=resWorkBook)
ws = resWorkBook.create_sheet('protein')

row = 2
column = 1
for key in expRes:
    ws.cell(row, column).value = key
    totalsum = 0
    preSum = 0
    for i in range(8):
        rowIndex = 0
        resStr = ''
        totalsum += len(expRes[key][i])
        expSum = len(expRes[key][i])
        if expSum != 0:
            for j in range(expSum):
                resStr += expRes[key][i][j]
                resStr += '\\'
            # 写rna名称
            ws.cell(row + rowIndex, column + 3).value = resStr

            # 写 三联体实验位点的 不同模式以及出现次数
            printResultXls(row, rowIndex, column, i, key, red)
            ws.cell(row + rowIndex + 1, column + 4).value = expSum

            # ws.merge_cells(start_row=row+rowIndex+1,start_column=column+4,end_row=row+rowIndex+1,end_column=column+6)
            # 写 三联体预测位点的不同模式、对应的rna名称以及出现次数
            for m in range(8):
                colIndex = 3
                if preRes[key][i][m] != 0:
                    # 写 预测位点模式
                    printResultXls(row, rowIndex, column + colIndex, i, key, green)
                    # 写 出现次数
                    ws.cell(row + rowIndex + 1, column + 4 + colIndex).value = preRes[key][i][m]
                    if preRes[key][i][m] == expSum:
                        break
                    else:
                        # 写rna名称
                        ws.cell(row + rowIndex + 1, column + 5 + colIndex).value = '\\'.join(
                            preRes[key][i][m + 8:m + 8 + preRes[key][i][m]])
                colIndex += 3

            rowIndex += 1

    # 写三联体出现次数
    ws.cell(row, column + 1).value = totalsum
    # 写带实验位点的三联体出现次数
    ws.cell(row, column + 2).value = totalsum - len(expRes[key][0])

    column += 1

print('hi')
resWorkBook.save(filePath + 'test.xlsx')

print('hello')