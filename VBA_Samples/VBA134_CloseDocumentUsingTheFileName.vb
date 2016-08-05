Sub CloseDocumentUsingTheFileName()
'VBA134
    Dim strFile As String
    strFile = Application.ApplicationDataPath & "test.pdl"
    Application.Documents.Close (strFile)
End Sub
