Sub CopySelectionToNewDocument()
'VBA138
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim iNewDoc As Integer
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    With objCircle
        .Top = 40
        .Left = 40
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 80
        .Selected = True
    End With
    MsgBox "Objects selected!"
    'Instead of "ActiveDocument.CopySelection" you can also write:
    '"ActiveDocument.Selection.CopySelection".
    ActiveDocument.CopySelection
    Application.Documents.Add hmiDocumentTypeVisible
    iNewDoc = Application.Documents.Count
    Application.Documents(iNewDoc).PasteClipboard
End Sub
