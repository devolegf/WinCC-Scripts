Sub EditTextList()
'VBA346
    Dim objTextList As HMITextList
    Set objTextList = ActiveDocument.HMIObjects("Textlist1")
    objTextList.BorderColor = RGB(255, 0, 0)
End Sub
