Sub ShowFirstObjectOfCollection()
'VBA232
    Dim strName As String
    strName = ActiveDocument.Application.AvailableDataLanguages(1).LanguageName
    MsgBox strName
End Sub
