import win32com.client
import win32com.client.gencache


Excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
wb = Excel.ActiveWorkbook
sheet = wb.ActiveSheet

# Excel.ActiveCell.SendKeys("ESCAPE")
wb1 = Excel.Workbooks.Add()
wb1.Activate()
wb.Activate()
# Excel.ActiveCell.Offset(2, 1).Activate()