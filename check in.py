from xlrd import open_workbook

workbook = open_workbook('number.xls')
worksheet = workbook.sheet_by_index(0)
n = int(worksheet.cell_value(0,0))
while(True):
    num = int(input("请输入门票码："))
    for i in range(1,n):
        x = worksheet.cell_value(i,4)
        if x == num:
            print("姓名：", worksheet.cell_value(i, 0))
            print("胸卡号：", int(worksheet.cell_value(i, 1)))
            print("学校：", worksheet.cell_value(i, 3))
            print("\n")
