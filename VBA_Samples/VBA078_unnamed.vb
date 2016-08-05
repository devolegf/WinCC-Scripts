Private Sub objGDApplication_BeforeDocumentClose(ByVal Document As IHMIDocument, Cancel As Boolean)
'VBA78
    MsgBox "The document " & Document.Name & " will be closed after press ok"
End Sub