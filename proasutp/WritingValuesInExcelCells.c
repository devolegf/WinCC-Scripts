'       (с) 2015 http://www.proasutp.com
'
'       Процедура-пример работы с ячейками таблицы Excel в WinCC на VBScript.
'
'       Вход:
'                Нет
'        Выход:
'                Нет

Sub WritingValuesInExcelCells ()

    Dim fso, myfile, i
    Dim objexcelapp

    Set fso = CreateObject("scripting.filesystemobject")
    Set myfile = fso.GetFile("d:\demo.xlsx")
    Set objexcelapp = CreateObject("excel.application")
   
    objexcelapp.visible=True
    objexcelapp.workbooks.open myfile
   
    i = objexcelapp.worksheets("sheet1").cells(2,100).value
    objexcelapp.worksheets("sheet1").cells(i,3).value = HMIRuntime.Tags("tag1").Read
    i = i + 1
    objexcelapp.worksheets("sheet1").cells(2,100).value = i
   
    objexcelapp.activeworkbook.Save
   
    objexcelapp.workbooks.close
    objexcelapp.quit
    Set objexcelapp = Nothing

End Sub