

'''https:#club.directum.ru/post/778?ysclid=l6djdl35ao769763968'''

import win32com.client
import win32com.client.gencache
from rich import print
from rich import inspect

# Word = CreateObject("Word.Application") 
Word = win32com.client.Dispatch("Word.Application")
Word = win32com.client.gencache.EnsureDispatch("Word.Application")
Doc = Word.Documents.Open(Path)
Doc = Word.ActiveDocument

'''Добавить текст'''
Doc.Paragraphs[2].Range.text = "12456"

'''Установка единицы измерения размера таблицы, 
где 3 - Сантиметры
    2 - проценты'''
Doc.Tables(1).PreferredWidthType = 2

'''Установка ширины таблицы'''
Doc.Tables(1).PreferredWidth = 100 # ширина таблицы в процентах
Doc.Tables(1).PreferredWidth = Word.CentimetersToPoints(17.5) # ширина таблицы в сантиметрах

'''Автоподбор размера ячейки таблицы
Фиксированная ширина = 0
По содержимому = 1
По ширине окна = 2
'''
Doc.Tables(1).AutoFitBehavior(2)

'''Поля в ячейках таблицы'''
tabWord = Doc.Tables(1)
tabWord.TopPadding = 0
tabWord.BottomPadding = 0
tabWord.LeftPadding = 0
tabWord.RightPadding = 0

'''Установка высоты ячеек'''
# Значение константы HeightRule
# Размер, указанный в параметре RowHeigh, является точным = 2
# Размер, указанный в параметре RowHeigh, является минимальным = 1
# Автоматический подбор высоты строк (параметр RowHeigh игнорируется) = 0

Doc.Tables(1).Rows.HeightRule = HeightRule   # указывает на способ изменения высоты
Doc.Tables(1).Rows.Height = RowHeigh         # RowHeight указывает на новую высоту строки в пунктах.

'''Установка интервала перед и после абзаца в таблице
Единица измерения интервала пт.
'''
Doc.Tables(1).Range.ParagraphFormat.SpaceBefore = 6  #интервал перед
Doc.Tables(1).Range.ParagraphFormat.SpaceAfter = 6   #интервал после

'''Абзац отступ слева'''
Doc.Tables(1).Range.ParagraphFormat.LeftIndent = 0.1 * 28.34646
tabWord.Cell(4, 2).Range.ParagraphFormat.LeftIndent = 0.1 * 28.34646
'''Отступ первой строки'''
Doc.Tables(1).Range.ParagraphFormat.FirstLineIndent = 1.2 * 28.34646
tabWord.Cell(4, 2).Range.ParagraphFormat.FirstLineIndent = 1.2 * 28.34646

'''Установка интервалов перед и после абзаца
'''
Doc.Paragraphs(1).Format.SpaceBefore = 12 # интервал перед
Doc.Paragraphs(1).Format.SpaceAfter = 12  # интервал после

'''Межстрочный интервал
0.5 - одинарный интервал
1 – полуторный интервал и 1.5 – двойной интервал.
'''
Doc.Paragraphs(1).Format.LineSpacingRule = 0.5 # одинарный интервал

'''Установка стиля таблицы'''
Doc.Tables(1).Style = ИмяСтиляТаблицы

'''Установка отступа слева '''
Doc.Tables(1).Rows.LeftIndent = 0

'''Абзацный отступ (красная строка) абзаца'''
Doc.Paragraphs(1).Format.FirstLineIndent = Значение в пунктах

'''Выделяем колонку и делаем отступ в ячейках'''
tabWord.Columns(2).Select()
col = Doc.Application.Selection
col.Font.Color = 255
col.ParagraphFormat.LeftIndent = 0.1 * 28.34646

'''Для перевода сантиметров в пункты можно воспользоваться функцией CentimetersToPoints, 
тогда абзацный отступ в 1,5 см можно задать следующим образом:'''
Doc.Paragraphs(1).Format.FirstLineIndent = Word.CentimetersToPoints(1.5)

'''Установка левой и правой границ текста абзаца:'''
Doc.Paragraphs(1).Format.LeftIndent = 10    # отступ слева
Doc.Paragraphs(1).Format.RightIndent = 10   # отступ справа

'''Установка значений полей ячеек по умолчанию'''
Doc.Tables(1).TopPadding = 0     # верхнее
Doc.Tables(1).BottomPadding = 0  # нижнее
Doc.Tables(1).LeftPadding = 0    # левое      
Doc.Tables(1).RightPadding = 0   # правое

'''Выравнивание текста в таблице по горизонтали
по левому краю = 0
по центру = 1
по правому краю = 2
по ширине = 3
'''
Doc.Tables(1).Range.ParagraphFormat.Alignment = 3  # по ширине
# в ячейке
Doc.Tables(1).Cell(2, 2).Range.ParagraphFormat.Alignment = 2


'''Выравнивание текстового абзаца:
Значения констант выравнивания такие же как в предыдущем пункте.
'''
Doc.Paragraphs(1).Format.Alignment = 0  # выравнивание по левому краю


'''Выравнивание текста в ячейке по вертикали
по верхнему краю = 0
по центру = 1
по нижнему краю = 3
'''
'''Выравниваем по вертикали ячейку'''
Doc.Tables(1).Rows(4).Cells(5).VerticalAlignment = 1
'''Выравниваем по вертикали строку'''
Doc.Tables(1).Rows(4).Cells.VerticalAlignment = 1
'''Выравниваем по вертикали столбец'''
Doc.Tables(1).Columns(4).Cells.VerticalAlignment = 1
'''Выравниваем по вертикали все ячейки в таблице'''
Doc.Tables(1).Range.Cells.VerticalAlignment = 1

'''Выбираем ячейку по номеру колонки и номера строки'''
fff = tabWord.Columns(2).Cells(4).Range
fff.Font.Color = 255



'''Установка размера шрифта таблицы'''
Doc.Tables(1).Range.Font.Size = 7

'''Установка цвета текста в ячейке'''
Doc.Tables(1).Cell(НомерСтроки, НомерСтолбца).Range.Font.Color = 255 

'''Выделение всего текста таблицы (жирным, курсивом, подчеркиванием)'''
Doc.Tables(1).Range.Font.Bold = True          # жирным
Doc.Tables(1).Range.Font.Italic = True        # курсивом
Doc.Tables(1).Range.Font.Underline  = True    # подчеркивание

'''Выделение или работа со строчкой по ячейке в ней'''
sheet.Cells(StartRow, col1).EntireRow.Delete()


'''Установка цвета подчеркивания'''
Doc.Tables(1).Cell(НомерСтроки, НомерСтолбца).Range.Font.UnderlineColor = 255

'''Установка темы шрифта таблицы'''
Doc.Tables(1).Range.Font.Name = "Arial"

'''Объединение ячеек'''
# # объединение первой и второй ячеек первой строки
Cell = Doc.Tables(1).Cell(1, 1)
Cell.Merge(Doc.Tables(1).Cell(1, 2))
'''Ячейка Текст жирный'''
cell3 = Doc.Tables(1).Cell(6, 3).Range.Font.Bold = True
'''Ячейка Текст курсив'''
cell3 = Doc.Tables(1).Cell(6, 3).Range.Font.Italic = False
'''Ячейка Текст масштаб по горизонтали (расстояние между буквами в слове)'''
Doc.Tables(1).Cell(6, 3).Range.Font.Scaling
'''Ячейка Цвет шрифта'''
Doc.Tables(1).Cell(6, 3).Range.Font.TextColor 
'''Ячейка Цвет подчеркнутой линии шрифта'''
Doc.Tables(1).Cell(6, 3).Range.Font.UnderlineColor  = 255
'''Ячейка подчеркнутая линии шрифта'''
Doc.Tables(1).Cell(6, 3).Range.Font.Underline = True


'''Вставка Excel таблицы в Word
Paragraph – номер параграфа, куда будет вставлена таблица из Excel.
'''
SelectionWord = Doc.Paragraphs(Paragraph).Range
SelectionWord.PasteExcelTable(True, False, False)

'''Разрыв связь'''
Doc.Fields.Unlink  

'''Удаление абзаца
Paragraph – номер параграфа, который нужно удалить.
'''
Doc.Paragraphs(Paragraph).Range.Delete

'''Удаляем строку в таблице'''
Doc.Tables(1).Rows(НомерСтроки).Select
Doc.Tables(1).Rows(НомерСтроки).Delete

'''Установка границ таблицы'''
# Линия, обрамляющая диапазон сверху = -1
# Линия, обрамляющая диапазон слева = -2
# Линия, обрамляющая диапазон снизу = -3
# Линия, обрамляющая диапазон справа = -4
# Все горизонтальные линии внутри диапазона = -5
# Все вертикальные линии внутри диапазона = -6
# Линия по диагонали сверху – вниз = -7
# Линия по диагонали снизу – вверх = -8

Table = Doc.Tables(1)
Table.Borders(WdBorderType).LineStyle = 4

'''Закрасить всю таблицу цветом'''
Doc.Tables(1).Shading.BackgroundPatternColor = 255 # заливка красным цветом

'''Закрасить ячейку цветом'''
Cell = Doc.Tables(1).Cell(1, 1)
Cell.Shading.BackgroundPatternColor = -687800525 # заливка желтым цветом

'''Установка ориентации страницы'''
Doc.Application.Selection.PageSetup.Orientation = 1 # альбомная
Doc.Application.Selection.PageSetup.Orientation = 0 # книжная

'''Установка полей страницы'''
Word.Application.Selection.PageSetup.LeftMargin = Word.CentimetersToPoints(2)    # левое поле
Word.Application.Selection.PageSetup.RightMargin = Word.CentimetersToPoints(2)   # правое поле
Word.Application.Selection.PageSetup.TopMargin = Word.CentimetersToPoints(2)     # верхнее поле
Word.Application.Selection.PageSetup.BottomMargin = Word.CentimetersToPoints(2)  # нижнее поле 

'''Добавление таблицы в документ'''
Doc.Tables.Add(Doc.Paragraphs(1).Range, 3, 5) # добавление таблицы из 5 столбцов и 3 строк в 1 абзац

'''Добавление строки в таблицу'''
Doc.Tables(1).Rows.Add

'''Добавление колонки в таблицу'''
Doc.Tables(1).Columns.Add

'''Добавление текста в ячейку'''
Doc.Tables(1).Cell(1,3).Range.Text = 'Текст, который добавляется в ячейку'

'''Сохранение документа в pdf формат:
где:
    Path - полный путь и имя нового файла формата PDF,
    17 - значение Microsoft.Office.Interop.Word.WdExportFormat, указывающие, что сохранять документ в формате PDF,
    openAfterExport - Значение True используется, чтобы автоматически открыть новый файл, 
        в противном случае используется значение False.
    CreateBookmarks - значение указывает, следует ли экспортировать закладки и тип закладки. 
        Значение константы WdExportCreateBookmarks:
    wdExportCreateHeadingBookmarks = 1 - Создание закладки в экспортируемом документе для всех заголовком, 
        которые включают только заголовки внутри основного документа и текстовые поля не в пределах колонтитулов, 
        концевых сносок, сносок или комментариев.
    wdExportCreateNoBookmarks = 0 - Не создавать закладки в экспортируемом документе.
    wdExportCreateWordBookmarks = 2 - Создание закладки в экспортируемом документе для каждой закладки, 
        которая включает все закладки кроме тех, которые содержатся в верхнем и нижнем колонтитулах.
'''
Doc.ExportAsFixedFormat(Path, 17, openAfterExport, CreateBookmarks) 

'''Для выделения определенного текста в документе можно воспользоваться, примерно, следующим кодом:'''

Word = win32com.client.CreateObject("Word.Application")  
Doc = Word.Documents.Open('Путь к документу', True, True)
myRange = Doc.Content 
# Поиск текста для выделения
myRange.Find.Execute("Выделяемый текст", True)     
isFind = myRange.Find.Found 
while isFind:
    # Выделения текста цветом
    myRange.Font.ColorIndex = 3
    myRange.Find.Execute("Выделяемый текст", True)     
    isFind = myRange.Find.Found
# endwhile           
Word.Visible = True 


'''Свойство таблицы "Повторять как заголовок на каждой странице"'''
Table.Rows.HeadingFormat = False
Table.Rows(1).HeadingFormat = True

'''Пример для удаления после разрыва страницы (до разрыва находится 
Закладка "СписокРассылки") страницы без содержания:'''
Word = win32com.client.CreateObject("Word.Application")  
WordDocument = Word.Documents.Open('Путь к документу', True, True)
CountTabl = WordDocument.Tables.Count           
# Удаление двух таблиц 6-ой и 5-ой
WordDocument.Tables.Item(CountTabl - 1).Delete
WordDocument.Tables.Item(CountTabl - 1).Delete                       
# Удаляем последний параграф с данными
CountP = WordDocument.Paragraphs.Count           
WordDocument.Paragraphs(CountP).Range.Delete
# Удаляем строки после Закладки
if WordDocument.Bookmarks.Exists("СписокРассылки"):         
    WordDocument.Bookmarks("СписокРассылки").Range.Delete
    WordDocument.Bookmarks("СписокРассылки").Range.Delete
    WordDocument.Bookmarks("СписокРассылки").Range.Delete                           
#WordDocument.Bookmarks("\Line").Range.Delete           
#WordDocument.Paragraphs(CountP - 2).Range.Delete

'''Сохранение файла по имени'''
Doc.save('table.docx')

'''Сохранить как'''
Doc.SaveAs(FileName = "generated.docx")

'''Если файл изменен, то сохранить его'''
if Doc.Saved == False: Doc.Save()

'''Создать по шаблону по шаблону'''
fail = os.getcwd() + "\\Шаблон_печати_А4_альбом.dotx"
Doc = Word.Documents.Add(fail)


'''Выбор объектов'''
tabWord = Doc.Tables(1)
Selection = tabWord.Select()
Selection = Word.Selection.SelectCell

'''Обтекание таблиц'''
tabWord = Doc.Tables(1)
tabWord.Rows.WrapAroundText = True


'''Нижний колонтитул'''
Footers = Doc.Sections(1).Footers
FootersCount = Footers.Count
Footers_2_Tables = Footers(2).Range.Tables
FootersTabCount = Footers_2_Tables.Count
FootersTables_2 = Footers_2_Tables(1)
FootersTables_2.Range.Cells.VerticalAlignment = 1

'''Верхний колонтитул'''
Headers = Doc.Sections(1).Headers
HeadersCount = Headers.Count
HeadersTables = Headers(1).Range.Tables
HeadersTabCount = HeadersTables.Count
HeadersTab_1 = HeadersTables(1)



def handle_updateRequest(rect=QtCore.QRect(), dy=0):
    '''Изменение высоты plainTextEdit и окна'''
    for widgetX in widgetList:
        doc = widgetX.document()
        tb = doc.findBlockByNumber(doc.blockCount() - 1)
        h = widgetX.blockBoundingGeometry(tb).bottom() + 2 * doc.documentMargin()
        widgetX.setFixedHeight(h)

    eee = sum([i.height() for i in widgetList])
    ''' если бы было 4 элемента, то они бы были со следующими размерами: 25 + 25 + 25 + 60 = 135; 60 - высота удаленного 4го элемента'''
    xxx = 0 if eee <= 135 else eee - 135
    Form.resize(Form.minimumWidth(), Form.minimumHeight() + xxx)
    

widgetList = [ui.plainTextEdit_4, ui.plainTextEdit_5, ui.plainTextEdit_6]
for widget in widgetList:
    widget.updateRequest.connect(handle_updateRequest)
