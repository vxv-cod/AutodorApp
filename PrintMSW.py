import os
import win32com.client
import win32com.client.gencache
import imageZeroFon
from VBAExcel import *
import time
import VXVtranslittext
# from Autodor import sig

# from rich import print
# from rich import inspect

# os.system('CLS') 


def GO(sheetName, ui, Form):
    print('---------------------------------------------------------')
    progressBar = ui.progressBar_1
    '''Создаем COM объект Excel'''
    try:
        Excel = win32com.client.GetActiveObject('Excel.Application')
    except:
        Excel = win32com.client.Dispatch("Excel.Application")

    # Excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
    
    Excel.Visible = 1
    wb = Excel.ActiveWorkbook
    sheet = wb.Worksheets(sheetName)

    sig.signal_Probar.emit(progressBar, 10)
    '''--------------------------------------------'''
    '''Находим номера крайней строки и столбца в таблице Excel'''
    EndRow, EndCol = EndIndexRowCol(sheet)
    StartRow, StartCol = 1, 1
    '''Выбираем таблицу в Excel'''
    tabEx = sheet.Range(sheet.Cells(StartRow, StartCol), sheet.Cells(EndRow, EndCol))
    # '''Копируем выбранный диапозон из Excel'''
    # tabEx.Copy()
    sleep(1)
    sig.signal_Probar.emit(progressBar, 20)

    '''---------------------------------------------------------------'''
    """Создаем COM объект Word"""
    Word = win32com.client.Dispatch("Word.Application")
    # Word = win32com.client.gencache.EnsureDispatch("Word.Application")
    Word.Visible = 1
    """Добавляем документ по шаблону"""
    sleep(1)

    strPath = ui.plainTextEdit_10.toPlainText()
    
    if strPath != '':
        strPath = strPath
    if strPath == '':
        # strPath = os.getcwd()
        strPath = wb.FullName.split(wb.Name)[0][:-1]


    nameobjextproekt = ui.plainTextEdit_4.toPlainText()

    if nameobjextproekt == '':
        text = f'Укажите ШИФР_ОБЪЕКТА: без учета символов ("-В-01", "-С-01")'
        sig.signal_err.emit(Form, text)
        return

    NameVOR = nameobjextproekt + '-В-01'
    NameOborud = nameobjextproekt + '-С-01'
    

    """Добавляем документ по шаблону"""
    if sheetName == "ВОР_Итог":
        ShifrObj = NameVOR
        fail = os.getcwd() + "\\Шаблон_печати_А4_книжная.dotx"
    else:
        ShifrObj = NameOborud
        fail = os.getcwd() + "\\Шаблон_печати_А4_альбом.dotx"

    # Doc = Word.Documents.Open(fail)
    # Doc = Word.Documents.Add(fail)
    # sleep(1)
    # Word.Visible = 1

    '''Сохранить как'''
    # 3129/5069/2-Р-001.004.390-НВК-01-С-001
    nameobjextproekt = VXVtranslittext.GO(nameobjextproekt)
    failPath = f"{strPath}\\{nameobjextproekt}.xlsx"
    wb.SaveAs(failPath, CreateBackup=0)
    
    saveAsNameobjextproekt = VXVtranslittext.GO(ShifrObj)
    FileName = f"{strPath}\\{saveAsNameobjextproekt}-rC01.docx"

    '''Проверяем открыт ли одноименный файл *.docx и подключаемся к нему,
    иначе создаем новый файл'''
    prov = False
    for i in Word.Documents:
        if i.FullName == FileName:
            prov = True
            Doc = i
    if prov == False:
        Doc = Word.Documents.Add(fail)
    
    Doc.Activate()
    Doc.SaveAs(FileName)
    sleep(1)


    '''Выбираем все документе'''
    myRange = Doc.Range()
    '''Копируем выбранный диапозон из Excel'''
    tabEx.Copy()
    sleep(1)
    '''Вставляем скопированную таблицу в World'''
    myRange.PasteExcelTable(False, False, False)
    '''Подключаемся к таблице'''
    tabWord = Doc.Tables(1)
    '''Автоподбор размера таблицы по содержимому'''
    tabWord.AutoFitBehavior(1)
    sleep(0.1)
    tabWord.AutoFitBehavior(2)

    '''Поля в ячейках таблицы'''
    tabWord.TopPadding = 0
    tabWord.BottomPadding = 0
    tabWord.LeftPadding = 0.05
    tabWord.RightPadding = 0
    tabWord.Spacing = 0
    tabWord.AllowPageBreaks = True
    tabWord.AllowAutoFit = True

    # for e in range(3, tabWord.Rows.Count):
    #     tabWord.Cell(e, 2).Range.ParagraphFormat.LeftIndent = 0.1 * 28.34646
    
    '''Высота строк в таблице'''
    PT = 28.34646   # количество "пт" в см
    Doc.Tables(1).Rows.HeightRule = 1               # указывает на способ изменения высоты: минимальный
    Doc.Tables(1).Rows.Height = 0.8 * PT            # RowHeight указывает на новую высоту строки в пунктах.
    try:
        Doc.Tables(1).Rows(StartRow).Height = 1.0 * PT          # RowHeight указывает на новую высоту строки в пунктах.
        Doc.Tables(1).Rows(StartRow + 1).Height = 1.2 * PT      # RowHeight указывает на новую высоту строки в пунктах.
    except:
        pass
    '''Обтекание таблиц'''
    tabWord.Rows.WrapAroundText = True
    '''Выравниваем по вертикали все ячейки в таблице'''
    Doc.Tables(1).Range.Cells.VerticalAlignment = 1
    '''Удаляем интервал после абзаца во всей таблице'''
    Doc.Tables(1).Range.ParagraphFormat.SpaceBefore = 0  #интервал перед
    Doc.Tables(1).Range.ParagraphFormat.SpaceAfter = 0   #интервал после
    # Doc.Tables(1).Range.ParagraphFormat.LeftIndent = 0.1 * 28.34646

    '''Выделяем колонку и делаем отступ в ячейках'''
    # tabWord.Columns(2).Select()
    # col = Doc.Application.Selection
    # col.ParagraphFormat.LeftIndent = 0.1 * 28.34646
    # Doc.Range(0, 0).Select()
    
    '''---------------------------------------------------------------'''
    sig.signal_Probar.emit(progressBar, 40)
    # sleep(1)

    '''Коллекция всех нижних колонтитулов'''
    Footers = Doc.Sections(1).Footers
    '''Подключаемся к 1-ой таблице нижнего колонтитула на 1-ом листе'''
    FootersTables_1 = Footers(1).Range.Tables(1)
    '''Все ячейки в таблице выравниваем по вертикали по центру'''
    FootersTables_1.Range.Cells.VerticalAlignment = 1
    '''Подключаемся к 1-ой таблице нижнего колонтитула на 2-ом листе'''
    FootersTables_2 = Footers(2).Range.Tables(1)
    '''Все ячейки в таблице выравниваем по вертикали по центру'''
    FootersTables_2.Range.Cells.VerticalAlignment = 1

    '''Исходные данные'''
    # patchPod = r"C:\Users\vvkhomutskiy\Desktop\TEST\ХВВ1.png"
    '''Спецификация'''
    # patchPod = ui.plainTextEdit_7.toPlainText()
    
    
    doljList = []
    UserList = []
    for i in range(ui.tableWidget_1.rowCount()):
        dolj = eval(f'ui.tableWidget_1.item({i}, 0).text()')
        xxx = eval(f'ui.tableWidget_1.item({i}, 1).text()')
        doljList.append(dolj)
        if dolj == '':
            UserList.append('')
        else:            
            UserList.append(xxx)
    
    rowList = 6, 7, 8, 9, 10, 11
    sec = time.localtime(time.time())
    now = f'{str(sec.tm_mday).rjust(2, "0")}.{str(sec.tm_mon).rjust(2, "0")}.{str(sec.tm_year)[-2:]}'

    def insertCellText(table, row, col, text):
        '''Отправляем текст в ячейку таблицы'''
        table.Cell(row, col).Range.Text = text
        # table.Cell(row, col).FitText = True

    '''Отправляем данные в штамп на 1-ом листе'''
    for i in range(len(doljList)):
        insertCellText(FootersTables_2, rowList[i], 2, doljList[i])
        if doljList[i] != '':
            insertCellText(FootersTables_2, rowList[i], 3, UserList[i])
            if len(UserList[i]) > 10:
                FootersTables_2.Cell(rowList[i], 3).FitText = True
            insertCellText(FootersTables_2, rowList[i], 5, now)



    sig.signal_Probar.emit(progressBar, 45)

    text = ShifrObj
    '''Отправляем данные в штамп на 2-ом листе'''
    '''ШИФР_ОБЪЕКТА'''
    insertCellText(FootersTables_1, 1, 8, text)

    '''Отправляем данные в штамп на 1-ом листе'''
    '''ШИФР_ОБЪЕКТА'''
    insertCellText(FootersTables_2, 1, 8, text)
    '''НАИМЕНОВАНИЕ_ОБЪЕКТА'''
    text = ui.plainTextEdit_5.toPlainText()
    insertCellText(FootersTables_2, 3, 8, text)
    '''НАИМЕНОВАНИЕ_РАЗДЕЛА'''
    text = ui.plainTextEdit_6.toPlainText()
    insertCellText(FootersTables_2, 6, 6, text)
    '''Спецификация'''
    if sheetName == "ВОР_Итог":
        text = 'Ведомость объемов строительных и монтажных работ'
    else:
        text = 'Спецификация оборудования, изделий и материалов'
    insertCellText(FootersTables_2, 9, 6, text)
    
    
    '''Стадия'''
    text = ui.plainTextEdit_8.toPlainText()
    insertCellText(FootersTables_2, 7, 7, text)

    '''---------------------------------------------------------------'''
    '''Работа с подписями'''
    directory = str(ui.plainTextEdit_3.toPlainText())

    if directory == '':
        sig.signal_err.emit(Form, "Подписи не были вставлены в штамп.\nУкажите папку с подписями в формате *.jpg , *.png")
        if Doc.Saved == False: Doc.Save()
        return

    try:
        direct = os.listdir(directory)
    except FileNotFoundError:
        sig.signal_err.emit(Form, "Папка с подписями не найдена")
        return

    '''Собираем список с полным именем файлов в папке с подписями'''
    PatchFileList = []
    for filename in direct:
        FullName = os.path.join(directory, filename)
        if os.path.isfile(FullName):
            if ".png" in FullName:
                PatchFileList.append(FullName)
                continue
            if ".jpg" in FullName:
                PatchFileList.append(FullName)

    
    patchPod = []
    userErr = []
    '''Для каждого значения фамилии из таблицы'''
    for User in UserList:
        UserTrue = False
        '''если оно не равно '' '''
        if User != '':
            '''перебираем все полные пути файлов в папке'''
            for patchP in PatchFileList:
                '''если фамилия есть в адрессе файла'''
                if User in patchP:
                    '''обозначаем наличие фамилии в названиях файлов в папке'''
                    UserTrue = True
                    '''добавляем адресс файла для фамилии из таблицы'''
                    patchPod.append(patchP)
                    '''Производит переход за пределы объемлющего цикла (всей инструкции цикла 
                    на уровень "for User in UserList") при нахождении фамилии'''
                    break
                    
            '''Если наличие файла в папке не подтвердилось'''                    
            if UserTrue == False:
                '''Cписок не найденных фамилий'''
                userErr.append(f'"{User}"')
                patchPod.append('')
        if User == '':
            patchPod.append('')  

    '''Работа над ошибками'''
    if userErr != []:
        UserErrList = ', '.join(userErr)
        text = f"Не найдены картинки с фамилией {UserErrList} в папке: \n{directory}"
        sig.signal_err.emit(Form, text)


    # printTabconsole([UserList, rowList, patchPod], add_column = True)
    '''Вставляем картинки с подписями в штамп и делаем их преде текстом'''
    xxx = 50
    for i in range(len(patchPod)):
        xxx += 8
        sig.signal_Probar.emit(progressBar, xxx)
        if patchPod[i] != '':
            if '..png' in patchPod[i]:
                FileName = patchPod[i]
            else:
                FileName = imageZeroFon.GO(patchPod[i])
            img = FootersTables_2.Cell(rowList[i], 4).Range.InlineShapes.AddPicture(FileName = FileName, LinkToFile = False, SaveWithDocument = True)
            img.ConvertToShape().WrapFormat.Type = 3
            sleep(0.5)

    if Doc.Saved == False: Doc.Save()

    '''---------------------------------------------------------------'''

if __name__ == "__main__":
    import sys
    from Autodor import  app, ui, Form
    GO("Спецификация_Итог", ui, Form)
    sys.exit(app.exec_())