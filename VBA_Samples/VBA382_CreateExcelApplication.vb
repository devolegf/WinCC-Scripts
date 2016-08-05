Sub CreateExcelApplication()
'VBA382
'
    'Open Excel invisible
    Dim objExcelApp As New Excel.Application
    MsgBox objExcelApp
    'Delete the reference to Excel and close it
    Set objExcelApp = Nothing
End Sub
