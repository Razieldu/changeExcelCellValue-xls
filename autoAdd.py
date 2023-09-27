
# import xlrd
# import xlwt
# from datetime import datetime
# from xlwt import easyxf
# # 指定你的 .xls 文件路径

# file_path = 'UserData.xls'
# output_file_path = 'UpdatedUserData.xls'


# # # 打开 .xls 文件
# workbook = xlrd.open_workbook(file_path)
# print(workbook)
# updated_workbook = xlwt.Workbook()
# # 获取工作表数量
# num_sheets = workbook.nsheets
# target_sheet_name = 'UserData'

# # date_style = NamedStyle(name='date_style', number_format='YYYY/MM/DD')
# # 逐一处理每个工作表
# for sheet_index in range(num_sheets):
#     sheet = workbook.sheet_by_index(sheet_index)

#     # 获取工作表名称
#     sheet_name = sheet.name

#     if sheet_name == target_sheet_name:
#         updated_sheet = updated_workbook.add_sheet(sheet_name)    
#         for row_index in range(sheet.nrows):
#             for col_index in range(sheet.ncols):
#                 cell_value = sheet.cell_value(row_index, col_index)
#                 if col_index == 0 and row_index == 1:
#                     cell_value = 'JPG0009999'
#                 elif col_index == 1 and row_index == 1:
#                     cell_value = 'JPG0009999'
#                 elif col_index == 2 and row_index == 1:
#                     cell_value = '測試中'
#                 elif col_index == 3 and row_index == 1 and  cell_value!="":    
#                     cell_value = ''
#                 elif col_index == 6 and row_index==1:
#                     date_value = xlrd.xldate_as_tuple(cell_value, workbook.datemode)
#                     python_date = datetime(*date_value).date()
#                     formatted_date = python_date.strftime('%Y/%m/%d')
#                 if col_index == 6 and row_index == 1:
#                    updated_sheet.write(row_index, col_index, formatted_date)  
#                 else:          
#                    updated_sheet.write(row_index, col_index, cell_value)
#     else:
#         updated_sheet = updated_workbook.add_sheet(sheet_name)
#         for row_index in range(sheet.nrows):
#             for col_index in range(sheet.ncols):
#                 cell_value = sheet.cell_value(row_index, col_index)
#                 updated_sheet.write(row_index, col_index, cell_value)

# # 最后，保存更新后的 .xls 文件
# updated_workbook.save(output_file_path)

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

# 修改单元格值

sheet.Cells(2, 1).Value = accountAndPassword
sheet.Cells(2, 2).Value = accountAndPassword
sheet.Cells(2, 3).Value = name
# 保存工作簿并关闭Excel
workbook.Save()
excel.Quit()

