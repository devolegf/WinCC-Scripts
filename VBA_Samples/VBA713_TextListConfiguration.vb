Sub TextListConfiguration()
'VBA713
    Dim objTextList As HMITextList
    Set objTextList = ActiveDocument.HMIObjects.AddHMIObject("myTextList", "HMITextList")
    With objTextList
        .SelBGColor = RGB(255, 0, 0)
    End With
End Sub
