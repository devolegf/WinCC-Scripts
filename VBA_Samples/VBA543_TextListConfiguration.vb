Sub TextListConfiguration()
'VBA543
    Dim objTextList As HMITextList
    Set objTextList = ActiveDocument.HMIObjects.AddHMIObject("myTextList", "HMITextList")
    With objTextList
        .ItemBorderStyle = 1
        .ItemBorderColor = RGB(255, 255, 255)
    End With
End Sub
