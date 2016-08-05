Sub TextListConfiguration()
'VBA762
    Dim objTextList As HMITextList
    Set objTextList = ActiveDocument.HMIObjects.AddHMIObject("myTextList", "HMITextList")
    With objTextList
        .UnselBGColor = RGB(255, 0, 0)
        .UnselTextColor = RGB(0, 0, 0)
    End With
End Sub
