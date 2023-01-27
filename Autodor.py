import os
import sys
from time import sleep
import win32com.client
import win32com.client.gencache
import threading
# from pythoncom import CoInitializeEx as pythoncomCoInitializeEx
import traceback
from PyQt5 import QtCore, QtWidgets
import pickle
import VXVtranslittext

from okno_ui import Ui_Form
from vxv_tnnc_SQL_Pyton import Sql
from version import ver
from VBAExcel import *
import PrintMSW

# from rich import print
# from rich import inspect
# inspect(xxx, methods=True)
# inspect(xxx, all =True)
# from prettytable import PrettyTable
# os.system('CLS')

app = QtWidgets.QApplication(sys.argv)
Form = QtWidgets.QWidget()
ui = Ui_Form()
ui.setupUi(Form)
Form.show()

_translate = QtCore.QCoreApplication.translate
Title = 'АвтоВОР v. 1.0' + str(ver)
Form.setWindowTitle(_translate("Form", Title))

def NFt(cells, okrug):
    try:
        cells.NumberFormat = okrug
    except:
        cells.NumberFormat = okrug.replace('.', ',')

def VPO(wb, sheet):
    EndRow, EndCol = EndIndexRowCol(sheet)

    EndRow = EndRow - 1

    dataAll = []
    def dataX(sheet, StartRow, EndRow, dataAll, col):
        cel = importdata(sheet, StartRow, col, EndRow, col)
        data = [i if i != '' else 0 for i in cel]
        # print(f"data = {data}")
        dataAll.append(data)
        return dataAll

    dataAll = dataX(sheet, StartRow = 5, EndRow = EndRow, dataAll = dataAll, col = 5)
    dataAll = dataX(sheet, StartRow = 5, EndRow = EndRow, dataAll = dataAll, col = 8)
    dataAll = dataX(sheet, StartRow = 5, EndRow = EndRow, dataAll = dataAll, col = 12)
            
    sheet1 = wb.Worksheets("Конструкция дорожной одежды...")
    EndRow1 = EndIndexRowCol(sheet)[0]
    # sheet.Activate()    
    dataAll = dataX(sheet1, StartRow = 6, EndRow = EndRow1, dataAll = dataAll, col = 18)

    # printTabconsole(dataAll)

    cell = sheet.Range("D4").Formula
    PiKet = int(cell.split('+')[0])
    metri = []
    piketAll = []
    schet = PiKet 
    countI = PiKet * 1000 + float(dataAll[0][0])
    metri.append(countI)
    
    nomerPiketList = []
    nasip = []
    oboch = []
    podstil = []
    nasip.append(float(dataAll[1][0]))
    oboch.append(float(dataAll[2][0]))
    podstil.append(float(dataAll[3][0]))
    nasipAll = []
    obochAll = []
    podstilAll = []
    for i in range(1, len(dataAll[0])):
        if PiKet <= schet:
            countI += float(dataAll[0][i])
            metri.append(countI)
            nasip.append(float(dataAll[1][i]))
            oboch.append(float(dataAll[2][i]))
            podstil.append(float(dataAll[3][i]))
            PiKet = int(str(countI / 1000).split('.')[0])
            nomerPiketList.append(PiKet)
        if PiKet > schet:
            schet = PiKet
            piketAll.append(metri)
            nasipAll.append(nasip)
            obochAll.append(oboch)
            podstilAll.append(podstil)
            nomerPiketList.append(PiKet)
            metri = []
            nasip = []
            oboch = []
            podstil = []
            i = i - 1

    PiketList = list(set(nomerPiketList))

    piketAll.append(metri)
    nasipAll.append(nasip)
    obochAll.append(oboch)
    podstilAll.append(podstil)
    
    nasipres = []
    obochres = []
    podstilres = []
    for i in range(len(PiketList)):
        nasipres.append(round(sum(nasipAll[i]), 2))
        obochres.append(round(sum(obochAll[i]), 2))
        podstilres.append(round(sum(podstilAll[i]), 2))

    return nasipres, obochres, piketAll, podstilres

def vvod(xxx):

    if xxx != '': 
        xxx = xxx.replace(',', '.')
    else: 
        xxx = 0
    try:
        xxx = round(float(xxx), 2)
    except:
        pass
    return xxx


def Pkzm():
    global Excel
    
    nameobjextproekt = ui.plainTextEdit_4.toPlainText()
    if nameobjextproekt == '':
        text = f'Укажите ШИФР_ОБЪЕКТА: без учета символов ("-В-01", "-С-01")'
        sig.signal_err.emit(Form, text)
        return
    
    delta = vvod((ui.plainTextEdit.toPlainText()))
    if isinstance(delta, str) or delta == 0:
        text = f'Расстояние от карьера задано не корректно или раыно нулю'
        sig.signal_err.emit(Form, text)
        return
    
    Kypl = vvod(ui.plainTextEdit_2.toPlainText())
    if isinstance(Kypl, str) or Kypl == 0:
        text = f'Коэффициент относительного уплотнения Купл задан не корректно . . .'
        sig.signal_err.emit(Form, text)
        return        

    Excel = win32com.client.gencache.EnsureDispatch('Excel.Application')

    # Excel.DisplayAlerts = False
    Excel.Visible = 1
    CountBook = Excel.Workbooks.Count

    wb1 = None
    if CountBook != 0:
        for i in range(1, CountBook + 1):
            Namebook = Excel.Workbooks(i).Name
            
            if "ведомость площадей и объемов" in Namebook:
                wb1 = Excel.Workbooks(Namebook)

    if wb1 == None:
        text = f"Подключение к файлу исходных данных \"Robur\" не состоялось, возможно файл не открыт . . ."
        sig.signal_err.emit(Form, text)

        # fail="АД от автодор с тв покр до куста 28 - ведомость площадей и объемов.xls"
        # wb1 = Excel.Workbooks.Open(os.getcwd() + rf"\{fail}")
        return

    if wb1 != None:
        sig.signal_Probar.emit(ui.progressBar_1, 20)
        Excel.Visible = 1
        # Excel.DisplayAlerts = False
        Excel.DisplayAlerts = True
        fail = os.getcwd() + "\\Автомобильные дороги.xltx"
        # wb, sheet = Book(fail=fail)
        wb = Excel.Workbooks.Open(fail)

        '''Сохраняем шаблон Excel под именем из ШИФР_ОБЪЕКТА'''
        strPath = ui.plainTextEdit_10.toPlainText()
        nameobjextproekt = ui.plainTextEdit_4.toPlainText()
        nameobjextproekt = VXVtranslittext.GO(nameobjextproekt)

        if strPath != '':
            strPath = strPath
        if strPath == '':
            strPath = wb1.FullName.split(wb1.Name)[0][:-1]
        wb.SaveAs(f"{strPath}\\{nameobjextproekt}.xlsx", CreateBackup=0)

        for i in ["Земляные работы...", "План. и укреп. работы...", "Конструкция дорожной одежды..."]:
            # wb.Activate()
            wb1.Worksheets(i).Copy(After=wb.Worksheets[wb.Worksheets.Count])
        
        '''Закрыть файл без сохранения'''
        # sleep(1)
        wb1.Close(False)
        sleep(1)
        
        Namebook = wb.Name

        sheet = wb.Worksheets("Земляные работы...")
        sheet.Activate()
        nasipres, obochres, piketAll, podstilres = VPO(wb, sheet)
        grunt = [nasipres[i] + obochres[i] + podstilres[i] for i in range(len(obochres))]
        nnn = len(grunt)

        # nnn_1 = nnn - 1
        F1 = f"=SUM(RC[-{nnn}]:RC[-1])"
        F2 = "=SUM(R[-3]C:R[-1]C)"
        F3 = "=R[-1]C"
        F4 = f"=R[-1]C*{Kypl}"
        F5 = "=R[-1]C*1.01"
        F6 = f"=SUM(R[-{nnn}]C:R[-1]C)"


        Shapka = [['Километры'] + [''] * nnn] + [[i for i in range(1, nnn + 1)] + ['Итого']]
        grunt.append(F1)
        data = Shapka + [grunt] + [[""] * nnn + [F1]] * 2 + [[F2] * (nnn) + [F1]]
        data = data + [[F3] * nnn + [F1]]
        data = data + [[F4] * nnn + [F1]]
        data = data + [[F5] * nnn + [F1]]
        data = data + [["Карьер №"] + [""] * nnn]
        
        sig.signal_Probar.emit(ui.progressBar_1, 30)
        for i in range(nnn):
            qqq = nnn - 1 - i
            eee = 2 + i
            data = data + [["-"] * (i) + [f"=R[-{eee}]C"] + ["-"] * (qqq) + [F1]]
        data = data + [[F6] * nnn + [F1]]

        ggg = "=("
        for i in range(nnn):
            qqq = nnn - i + 1
            xxx = f"R[-{qqq}]C[{nnn}]*R[-{qqq}]C[-1]"
            ggg += xxx + "+"
        xxx = ggg[:-1] + f")/R[-1]C[{nnn}]"
        data = data + [[xxx] + [""] * (nnn)]
        
        # printTabconsole(data, [str(i) for i in range(1, nnn + 1)] + ["Итого"])

        sheet = wb.Worksheets("ПКЗМ")
        sheet.Activate()
        
        xxx = str(Kypl).replace('.', ',')
        textB9 = f'Итого насыпи из песка с учетом коэффициента относительного уплотнения Купл={xxx}'
        cel = sheet.Range("B9").Formula = textB9
        
        sig.signal_Probar.emit(ui.progressBar_1, 40)
        StartRow = 2
        EndRow = StartRow + len(data) - 1
        StartCol = 4
        EndCol = StartCol + nnn
        StartCol111 = StartCol
        EndCol111 = EndCol
    
        exportdata(data, sheet, StartRow, StartCol, EndRow, EndCol)
        sig.signal_Probar.emit(ui.progressBar_1, 60)

        fff = [["", ""] + [delta + piketAll[i][-1]/1000] for i in range(nnn)]

        textAList111 = [["Наименование месторождения грунта", "", ""]]
        textAList222 = [["Погрузка грунта (песок) экскаватором в автомобили самосвалы и транспортировка в тело насыпи на расстояние до, км", "", f"{delta}"]] + fff[:-1]

        fff = [
            ["Итого по карьеру", "", ""],
            ["Средневзвешенная дальность возки, км", "", ""],
            ["Примечание:", "", ""],
            ["1. Дополнительные объемы включают в себя объемы на устройство примыканий, углов поворота, берм для установки дорожных знаков, обратной засыпки котлованов у водопропускных труб;", "", ""],
            ["2. В объеме насыпи ниже дневной поверхности учтены поправки на осадку земляного полотна на слабом основании и сжатие почвенно-растительного слоя.", "", ""]
            ]

        textAList = textAList111 + textAList222 + fff
        # printTabconsole(textAList, align = "l", column = 1)
        sig.signal_Probar.emit(ui.progressBar_1, 70)
        
        data = textAList
        StartRow = 11
        StartCol = 1
        EndRow = StartRow + len(data) - 1
        EndCol = StartCol + len(data[0]) - 1        
        exportdata(data, sheet, StartRow, StartCol, EndRow, EndCol)
        
        sig.signal_Probar.emit(ui.progressBar_1, 80)

        cel = sheet.Range(sheet.Cells(2, StartCol111), sheet.Cells(2, EndCol111)).Merge()
        cel = sheet.Range(sheet.Cells(11, 1), sheet.Cells(11, StartCol111 - 1)).Merge()
        '''Наименование месторождения грунта'''
        cel = sheet.Range(sheet.Cells(11, StartCol111), sheet.Cells(11, EndCol111))
        cel.Merge()
        cel.RowHeight = 20
        '''Погрузка грунта (песок) экскаватором в автомобили самосвалы и 
        транспортировка в тело насыпи на расстояние до, км'''
        StartRow = 12
        EndRow = StartRow + nnn - 1
        cel = sheet.Range(sheet.Cells(StartRow, 1), sheet.Cells(EndRow, 2))
        cel.Merge()
        cel.RowHeight = 21
        '''Итого по карьеру'''
        StartRow = EndRow + 1
        cel = sheet.Range(sheet.Cells(StartRow, 1), sheet.Cells(StartRow, 3))
        cel.Merge()
        cel.RowHeight = 15
        '''Средневзвешенная дальность возки, км'''
        StartRow = StartRow + 1
        EndCol = 3
        cel = sheet.Range(sheet.Cells(StartRow, 1), sheet.Cells(StartRow, EndCol))
        cel.Merge()
        cel.RowHeight = 15
        cel = sheet.Range(sheet.Cells(StartRow, 4), sheet.Cells(StartRow, EndCol111))
        cel.Merge()
        sig.signal_Probar.emit(ui.progressBar_1, 90)
        '''Примечание:'''
        EndCol = 4 + nnn
        StartRow = StartRow + 1
        cel = sheet.Range(sheet.Cells(StartRow, 1), sheet.Cells(StartRow, EndCol))
        cel.Merge()
        '''1. Дополнительные объемы включают в себя'''
        cel.RowHeight = 15
        StartRow = StartRow + 1
        cel = sheet.Range(sheet.Cells(StartRow, 1), sheet.Cells(StartRow, EndCol))
        cel.Merge()
        cel.RowHeight = 35
        '''2. В объеме насыпи ниже дневной поверхности '''
        StartRow = StartRow + 1
        cel = sheet.Range(sheet.Cells(StartRow, 1), sheet.Cells(StartRow, EndCol))
        cel.Merge()
        cel.RowHeight = 35

        EndRow = StartRow
        cel = RangeCells(sheet, 2, 1, EndRow - 3, EndCol)
        cel.Borders.Weight = 2

        cel = RangeCells(sheet, EndRow - 4, 1, EndRow, EndCol)
        cel.Borders(7).Weight = 2
        cel.Borders(8).Weight = 2
        cel.Borders(9).Weight = 2
        cel.Borders(10).Weight = 2
    sleep(1)


def VorShablonCorrect():
    Kypl = vvod(ui.plainTextEdit_2.toPlainText())
    if isinstance(Kypl, str) or Kypl == 0:
        text = f'Коэффициент относительного уплотнения Купл задан не корректно . . .'
        sig.signal_err.emit(Form, text)
        return  

    Excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
    wb = Excel.ActiveWorkbook
    sheet = wb.Worksheets("ПКЗМ")
    sheet.Activate()
    NameEndColumn, NameEndRow = NameEndCell(sheet)
    End8 = f"=ПКЗМ!{NameEndColumn}8"
    Itogo = f"=ПКЗМ!{NameEndColumn}{int(NameEndRow) - 4}"
    # srdv = f"=ПКЗМ!D{int(NameEndRow) - 4}"
    srdv = round(sheet.Range(f"D{int(NameEndRow) - 3}").Value)

    sig.signal_Probar.emit(ui.progressBar_1, 30)
    

    sheet = wb.Worksheets("План. и укреп. работы...")
    sheet.Activate()
    NameEndColumn, NameEndRow = NameEndCell(sheet)
    OtkosiNasipi = rf"='План. и укреп. работы...'!G{NameEndRow}"

    sheet = wb.Worksheets("ВОР_шаблон")
    sheet.Activate()
    
    sig.signal_Probar.emit(ui.progressBar_1, 40)
    cel = sheet.Range("D30")
    cel.Formula = End8
    cel.Font.Color = -4165632
    NFt(cel, "0")

    sig.signal_Probar.emit(ui.progressBar_1, 50)
    cel = sheet.Range("D26")
    cel.Formula = Itogo
    cel.Font.Color = -4165632
    NFt(cel, "0")
    
    sig.signal_Probar.emit(ui.progressBar_1, 60)
    cel = sheet.Range("B26")
    textB26 = f'Разработка грунта 1 группы (песок) в карьере № … экскаватором с погрузкой в автосамосвалы и транспортировкой в насыпь на расстояние до {srdv} км (Кпот = 1,01, Купл = {Kypl}, γгр. = … т/м3)'
    cel.Formula = textB26
    NFt(cel, "0")

    sig.signal_Probar.emit(ui.progressBar_1, 75)
    cel = sheet.Range("D33")
    cel.Formula = OtkosiNasipi
    cel.Font.Color = -4165632
    NFt(cel, "0")
    
    sig.signal_Probar.emit(ui.progressBar_1, 90)
    cel = sheet.Range("D92")
    cel.Formula = OtkosiNasipi
    cel.Font.Color = -4165632
    NFt(cel, "0")
    
    sheet.Range("A1").Select()
    sleep(1)


def VOR(sheetCopy, sheetPaste, col1, col2, col3, col4):
    print('---------------------------------------------------------')
    global proc
    Excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
    # Excel = win32com.client.Dispatch('Excel.Application')
    Namebook = Excel.Workbooks(Excel.Workbooks.Count).Name
    if Namebook:
        wb = Excel.Workbooks(Namebook)

        """Получаем доступ к определенному листу"""
        sheet = wb.Worksheets(sheetCopy)
        sheet.Activate()
        EndRow, EndCol = EndIndexRowCol(sheet)
        StartRow = 1
        
        '''Копируем ячейки'''
        cel = sheet.Cells
        cel.Copy()

        sheet = wb.Worksheets(sheetPaste)
        sheet.Activate()
        sheet.Range("A1").Select()
        sheet.Paste()
        uuu = 15
        proc += uuu
        sig.signal_Probar.emit(ui.progressBar_1, proc)
                
        
        def cellsSelect():
            cel_1 = sheet.Range(sheet.Cells(StartRow, col1), sheet.Cells(EndRow, col1))
            cel_2 = sheet.Range(sheet.Cells(StartRow, col2), sheet.Cells(EndRow, col2))
            cel_3 = sheet.Range(sheet.Cells(StartRow, col3), sheet.Cells(EndRow, col3))
            cel_4 = sheet.Range(sheet.Cells(StartRow, col4), sheet.Cells(EndRow, col4))
            return cel_1, cel_2, cel_3, cel_4
        
        cel_1, cel_2, cel_3, cel_4 = cellsSelect()

        for i in range(3, len(cel_4) + 1):
            '''Удаляем строки с пустым количествов'''
            while cel_3[i].Formula != "" and (cel_4[i].Formula == "" or cel_4[i].Formula == 0 or cel_4[i].Value == -2146826273 or cel_4[i].Value == -2146826265):
                # print(f"111 - {i} --> {cel_1[i].Formula} {cel_2[i].Formula}")
                cel_4[i].EntireRow.Delete()

        proc += uuu
        sig.signal_Probar.emit(ui.progressBar_1, proc)
        sleep(0.5)

        EndRow, EndCol = EndIndexRowCol(sheet)
        cel_1, cel_2, cel_3, cel_4 = cellsSelect()
        
        '''Удаляем строки с пустым ед. изм. и количеством с заполненной следующей строкой'''
        for i in range(3, len(cel_4) + 1):
            while cel_1[i].Formula != "" and cel_2[i].Formula != "" and cel_3[i].Formula == "" and cel_4[i].Formula == "" and cel_1[i+1].Formula != "":
                cel_4[i].EntireRow.Delete()
        
        '''Удаляем названия групп, если элементов нет'''
        if sheetPaste == "ВОР_Итог":
            for i in range(3, len(cel_4) + 1):
                while cel_2[i].Font.Bold == True and cel_2[i+1].Font.Bold == True:
                    cel_2[i].EntireRow.Delete()

        '''Удаляем строки с нулями'''
        if sheetPaste == "Спецификация_Итог":
            for i in range(3, len(cel_4) + 1):
                while cel_4[i].Value == 0:
                    cel_4[i].EntireRow.Delete()

       
        sheet.Range("A1").Select()
        
        proc += uuu
        sig.signal_Probar.emit(ui.progressBar_1, proc)
        
        '''Перебиваем сквознцю нумерацию'''
        nomer = 1
        for i in range(3, len(cel_1) + 1):
            xxx = cel_1[i].Formula
            if xxx != "":
                cel_1[i].Formula = nomer
                nomer += 1
    else:
        errortext = traceback.format_exc()
        text = f'Файл шаблона не найден, повторите попытку . . . \n\n{errortext}'
        sig.signal_err.emit(Form, text)
        return
    return wb
    

def startFun(my_func):
    '''Обертка функции (декоратор)'''
    def wrapper():
        Sql("Autodor")
        pushButtonList = [ui.pushButton, ui.pushButton_2, ui.pushButton_3, ui.pushButton_4]
        progressBar = ui.progressBar_1
        label = ui.label
        try:
            for i in pushButtonList:
                sig.signal_bool.emit(i, True)
            sig.signal_label.emit(label, "Обработка данных . . .")
            sig.signal_Probar.emit(progressBar, 0)
            sig.signal_color.emit(progressBar, 0)
            my_func()
        except:
            errortext = traceback.format_exc()
            print(errortext)
            text = f"Ошибка работы, повторите попытку \n\n{errortext}"
            sig.signal_err.emit(Form, text)
        for i in pushButtonList:
            sig.signal_bool.emit(i, False)
        sig.signal_Probar.emit(progressBar, 0)
        sig.signal_color.emit(progressBar, 100)
        sig.signal_label.emit(label, "Выполнено . . .")
    return wrapper

'''--------------------------------------------------------------------'''
'''Ручная простановка высоты формы'''
# Form.setMinimumHeight(500)

def handle_updateRequest(rect=QtCore.QRect(), dy=0):
    '''Изменение высоты plainTextEdit и окна'''
    for widgetX in widgetList:
        doc = widgetX.document()
        tb = doc.findBlockByNumber(doc.blockCount() - 1)
        h = widgetX.blockBoundingGeometry(tb).bottom() + 2 * doc.documentMargin()
        widgetX.setFixedHeight(h)

        eee = sum([widgetList[i].height() for i in range(len(widgetList) - 1)])
        ''' (25 * 3 = 75) + 86 = 161; 86 - максимальная высота последнего элемента с подставленным Vertical Spacer'''
        hhh = 3 * 25 + 86
        xxx = 0 if eee <= hhh else eee - hhh
        '''Сравниваем максимальные значения на разных вкладках'''
        rrr = widgetList[-1].height() - 25
        xxx = max(xxx, rrr)
    '''Корректируем высоту формы, если размеры widgetX больше допустимых'''
    Form.resize(Form.minimumWidth(), Form.minimumHeight() + xxx)
    
widgetList = [ui.plainTextEdit_4, ui.plainTextEdit_5, ui.plainTextEdit_6, ui.plainTextEdit_8, ui.plainTextEdit_3]
for widget in widgetList:
    widget.updateRequest.connect(handle_updateRequest)
'''--------------------------------------------------------------------'''

@thread
@startFun
def startPkzm():
    Pkzm()

@thread
@startFun
def startVorShablon():
    VorShablonCorrect()

@thread
@startFun
def startVOR():
    global proc
    proc = 0
    VOR("ВОР_шаблон", "ВОР_Итог", 1, 2, 3, 4)
    sleep(0.5)
    VOR("Спецификация_шаблон", "Спецификация_Итог", 1, 2, 6, 7)

@thread
@startFun
def startWorld():
    PrintMSW.GO("ВОР_Итог", ui, Form)
    PrintMSW.GO("Спецификация_Итог", ui, Form)
    

ui.pushButton.clicked.connect(startPkzm)
ui.pushButton_2.clicked.connect(startVorShablon)
ui.pushButton_3.clicked.connect(startVOR)
ui.pushButton_4.clicked.connect(startWorld)
'''Запускаем функцию с аргументами'''
ui.pushButton_5.clicked.connect(lambda  : redactExcel("Автомобильные дороги.xltx"))


'''Отслеживаем сигнал закрытия окна и сохраняем все из окна перед закрытием'''
def AppQuit():
    doljList = []
    UserList = []
    for i in range(ui.tableWidget_1.rowCount()):
        dolj = eval(f'ui.tableWidget_1.item({i}, 0).text()')
        xxx = eval(f'ui.tableWidget_1.item({i}, 1).text()')
        doljList.append(dolj)
        UserList.append(xxx)

    saveData = [
                ui.plainTextEdit_3.toPlainText(),
                doljList,
                UserList,
                ui.plainTextEdit_8.toPlainText(),
                ui.plainTextEdit_2.toPlainText()
                ]

    with open("saveData.ini", "wb") as f:
        pickle.dump(saveData, f) # помещаем объект в файл
app.aboutToQuit.connect(AppQuit)

'''Чистим "plainTextEdit" для отображения текста по умолчанию'''
ui.plainTextEdit.clear()
ui.plainTextEdit_2.clear()
ui.plainTextEdit_4.clear()
ui.plainTextEdit_5.clear()
ui.plainTextEdit_6.clear()
ui.plainTextEdit_8.clear()
ui.plainTextEdit_10.clear()

"""Если файл с данными НЕ существует"""
savePathFile = os.getcwd() + "\saveData.ini"
if os.path.exists(savePathFile) == False:
    with open("saveData.ini", "wb") as file:
        pass


'''Заполняем значения с данными в форму из файла после запуска программы'''
with open("saveData.ini", "rb") as f:
    try:
        loadx = pickle.load(f) # извлекаем ообъект из файла
        ui.plainTextEdit_3.setPlainText(f"{loadx[0]}")
        ui.plainTextEdit_8.setPlainText(f"{loadx[3]}")
        ui.plainTextEdit_2.setPlainText(f"{loadx[4]}")
        for i in range(ui.tableWidget_1.rowCount()):
            eval (f'ui.tableWidget_1.item({i}, {0}).setText(_translate("Form", "{loadx[1][i]}"))')
            eval (f'ui.tableWidget_1.item({i}, {1}).setText(_translate("Form", "{loadx[2][i]}"))')
    except:
        pass

'''Отслеживаем сигнал в plainTextEdit на изменение данных и удаляем не нужный текст'''
def ChangedPT(plainTextEdit):
    '''Удаления ненужного текста в plainTextEdit_3'''
    # ui.plainTextEdit_3.clear()
    directory = plainTextEdit.toPlainText()
    if "file:///" in directory:
        xxx = directory.rfind("file:///")
        directory = directory[xxx + 8:]
        try:
            directory = directory.replace("/", "\\")
        except:
            pass
        plainTextEdit.setPlainText(rf"{directory}")
ui.plainTextEdit_3.textChanged.connect(lambda : ChangedPT(ui.plainTextEdit_3))
ui.plainTextEdit_10.textChanged.connect(lambda : ChangedPT(ui.plainTextEdit_10))

'''--------------------------------------------------------------------'''

if __name__ == "__main__":
    # startPkzm()
    # startVOR()
    sys.exit(app.exec_())
    