import inspect
import win32com.client
import win32com.client.gencache
import os
from pythoncom import CoInitializeEx as pythoncomCoInitializeEx

# pythoncomCoInitializeEx(0)
# Excel = win32com.client.Dispatch("Excel.Application")
# Excel = win32com.client.gencache.EnsureDispatch('Excel.Application')

# fail="АД от автодор с тв покр до куста 28 - ведомость площадей и объемов.xls"
# wb1 = Excel.Workbooks.Open(os.getcwd() + rf"\{fail}")

# fail = os.getcwd() + "\\Автомобильные дороги.xlsx"
# wb = Excel.Workbooks.Open(fail)


# # sheet = wb.Worksheets('ВОР_шаблон')
# # sheet.Activate()
# # cel = sheet.Range("A1")
# # cel.Activate()
# # cel.Formula = 555

# # sheet.Activate()
# # wb1.Worksheets("Земляные работы...").Copy(After=wb.Worksheets[2])
# wb1.Worksheets("Земляные работы...").Copy(After=wb.Worksheets[2])

Word = win32com.client.Dispatch("Word.Application")
Doc = Word.ActiveDocument
tabWord = Doc.Tables(1)
tabWord.TopPadding = 0
tabWord.BottomPadding = 0
tabWord.LeftPadding = 0
tabWord.RightPadding = 0


# tabWord.Cell(4, 2).LeftPadding = 5
# tabWord.Columns(1).LeftPadding = 5

# PT = 28.34646
# tabWord.Range.ParagraphFormat.LeftIndent = 0.1 * 28.34646
# tabWord.Cell(4, 2).Range.ParagraphFormat.FirstLineIndent = 1.2 * 28.34646
# tabWord.Cell(4, 2).Column.Cells.ParagraphFormat.FirstLineIndent = 1.2 * 28.34646

# tabWord.Range(tabWord.Cell(3, 2).Start, tabWord.Cell(20, 2).End)
# .ParagraphFormat.LeftIndent = 0.5 * 28.34646
# for e in range(3, tabWord.Rows.Count):
#     tabWord.Cell(e, 2).Range.ParagraphFormat.LeftIndent = 0.5 * 28.34646


# rs = tabWord.Rows(3).Range.Start
# re = tabWord.Rows(7).Range.End
# fff = Doc.Range(rs, re)
# # fff.ParagraphFormat.LeftIndent = 0.2 * 28.34646
# fff.Font.Color = 255 

# fff = tabWord.Columns(2).Range
# fff = Doc.Tables(1).Cell(3, 2).Range.Columns.Select
# tabWord.Cell(3, 2).Select
# fff = Doc.Application.Selection.SelectColumn
# fff.Range.text = "12456"
# print(fff)

# fff = Doc.Range(tabWord.Cell(4, 2).Range.Start, tabWord.Cell(7, 2).Range.End).Select
# Doc.Application.Selection.Font.Color = 255 


# fff.Range.ParagraphFormat.LeftIndent = 1.1 * 28.34646

# fff = Doc.Range(tabWord.Cell(4, 2).Range.Start, tabWord.Cell(8, 2).Range.End)
# fff.Font.Color = 255 
# # for i in fff:
#     i.text = "aaaaaaaaaa"

# fff = tabWord.Columns(2).Cells(5).Range
# fff.Font.Color = 255 
# for i in range(5, 20):
#     fff = tabWord.Columns(2).Cells(i).Range
    # fff.Font.Color = 255 

# fff = tabWord.Columns(2).Cells
# fff.Font.Color = 255 

from rich import inspect
# inspect(tabWord.Columns(2).Select(), all=True)


tabWord.Columns(2).Select()
col = Doc.Application.Selection
col.Font.Color = 255
col.ParagraphFormat.LeftIndent = 0.1 * 28.34646

tabWord.AutoFitBehavior(1)
Doc.Range(0, 0).Select()


