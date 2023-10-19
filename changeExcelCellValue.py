import win32com.client
import pyautogui
import time
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
ifdelete= input("是否執行刪除")


sheet.Cells(2, 1).Value = accountAndPassword
sheet.Cells(2, 2).Value = accountAndPassword
sheet.Cells(2, 3).Value = name
# 保存工作簿并关闭Excel
workbook.Save()
excel.Quit()


pyautogui.hotkey('win', 'd')
pyautogui.moveTo(1540, 20) ###轉換的檔案擺在右上角
pyautogui.click()
pyautogui.hotkey('ctrl', 'c')
pyautogui.moveTo(420, 940) #遠端在下面數來第1個
pyautogui.click()
pyautogui.moveTo(120,280) #portal位置
pyautogui.doubleClick()


######remote視窗操作######
time.sleep(2)
##移動到叉叉位置關閉多餘視窗
pyautogui.moveTo(1582,104) 
pyautogui.click()
pyautogui.click()


##刪除操作(遠端桌面已經有檔案的時候執行此部分程式碼)
if ifdelete =="1":
 pyautogui.moveTo(1000,600)
 pyautogui.rightClick() 
 pyautogui.moveTo(1050,525)
 pyautogui.click()
 pyautogui.moveTo(1050,530)
 pyautogui.click()

###貼上操作
pyautogui.moveTo(1000,600)
time.sleep(1)
pyautogui.rightClick()
time.sleep(1) 
pyautogui.moveTo(1050,460)
pyautogui.click()
