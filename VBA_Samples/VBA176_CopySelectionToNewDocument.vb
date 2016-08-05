Sub CopySelectionToNewDocument()
'VBA176
    Dim iNewDoc As String
    ActiveDocument.CopySelection
    Application.Documents.Add hmiDocumentTypeVisible
    iNewDoc = Application.Documents.Count
    Application.Documents(iNewDoc).PasteClipboard
End Sub
