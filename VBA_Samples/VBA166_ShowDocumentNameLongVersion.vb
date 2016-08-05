Sub ShowDocumentNameLongVersion()
'VBA166
    Dim strDocName As String
    strDocName = Application.Documents.Item(3).Name
    MsgBox strDocName
End Sub
