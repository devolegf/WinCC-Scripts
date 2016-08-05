Sub ActiveDocumentConfiguration()
'VBA552
    Dim varLastDocChange As Variant
    varLastDocChange = Application.ActiveDocument.LastChange
    MsgBox "Last changing: " & varLastDocChange
End Sub
