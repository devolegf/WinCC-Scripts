Sub TextListConfiguration()
'VBA542
    Dim objTextList As HMITextList
    Set objTextList = ActiveDocument.HMIObjects.AddHMIObject("myTextList", "HMITextList")
    With objTextList
        .ItemBorderStyle = 1
        .ItemBorderBackColor = RGB(255, 0, 0)
    End With
End Sub
