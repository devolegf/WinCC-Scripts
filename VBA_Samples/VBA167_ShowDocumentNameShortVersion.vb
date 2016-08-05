Sub ShowDocumentNameShortVersion()
'VBA167
    Dim strDocName As String
    strDocName = Application.Documents(3).Name
    MsgBox strDocName
End Sub
