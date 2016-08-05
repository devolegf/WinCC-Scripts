Sub ShowDocuments()
'VBA482
    Dim colDocuments As Documents
    Dim objDocument As Document
    Dim strOutput As String
    Set colDocuments = Application.Documents
    strOutput = "List of all opened documents:" & vbCrLf
    For Each objDocument In colDocuments
        strOutput = strOutput & vbCrLf & objDocument.Name
    Next objDocument
    MsgBox strOutput
End Sub
