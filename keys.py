import xlwt,random

book = xlwt.Workbook(encoding='utf-8')
sheet = book.add_sheet('Keys')
sheet.write(0,0,1)
sheet.write(0,1,1)
sheet.write(0,2,1)
for x in range(1,301):
    sheet.write(x,4,random.randint(100000,999999))
book.save('number.xls')