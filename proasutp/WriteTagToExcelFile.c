'       (с) 2015 http://www.proasutp.com
'
'       Процедура-пример записи значения тега в таблицу Excel в WinCC на VBScript.
'
'       Вход:
'                Нет
'        Выход:
'                Нет

Sub WriteTagToExcelFile () 
    Dim fso, myfile
    Dim objexcelapp
    Dim path, filename
   
    Set fso = CreateObject ("scripting.filesystemobject")
    Set myfile = fso.GetFile ("c:\demo.xlsx")
   
    Set objexcelapp = CreateObject ("excel.application")
    objexcelapp.visible=True
   
    objexcelapp.workbooks.open myfile
   
    objexcelapp.worksheets("sheet1").cells(2,3).value = HMIRuntime.Tags("tag1").Read
   
    filename = CStr(Year(Now)) & "-" & CStr(Month(Now)) & "-" & CStr(Day(Now))
    filename = "-" & CStr(Hour(Now)) & "-" & CStr(Minute(Now)) & "-" & CStr(Second(Now))
    path = "c:\" & filename & "-" & "demo.xlsx"
   
    objexcelapp.activeworkbook.SaveAs path
   
    objexcelapp.workbooks.close
    objexcelapp.quit
    Set objexcelapp = Nothing
   
End Sub