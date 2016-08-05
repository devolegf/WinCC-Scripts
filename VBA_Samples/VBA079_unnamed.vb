Private Sub objGDApplication_BeforeDocumentSave(ByVal Document As IHMIDocument, Cancel As Boolean)
'VBA79
    MsgBox Document.Name & "-saving will start after press ok."
End Sub