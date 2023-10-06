import win32com.client

# 创建Excel应用程序对象
excel = win32com.client.Dispatch("Excel.Application")

# 打开工作簿
workbook = excel.Workbooks.Open(r'C:\Users\jeremydu\Desktop\UserData.xls')

# 获取工作表
sheet = workbook.Sheets("UserData")

# 读取单元格值
# value = sheet.Cells(1, 1).Value
accountAndPassword = input("輸入以作為帳號及密碼__ ")
name = input("輸入姓名__")

sheet.Cells(2, 1).Value = accountAndPassword
sheet.Cells(2, 2).Value = accountAndPassword
sheet.Cells(2, 3).Value = name
# 保存工作簿并关闭Excel
workbook.Save()
excel.Quit()