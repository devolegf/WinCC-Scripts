Sub ShowDocuments()
'VBA238
    Dim colDocuments As Documents
    Dim objDocument As Document
    Set colDocuments = Application.Documents
    For Each objDocument In colDocuments
        MsgBox objDocument.Name
    Next objDocument
End Sub
