Sub TextListConfiguration()
'VBA716
    Dim objTextList As HMITextList
    Set objTextList = ActiveDocument.HMIObjects.AddHMIObject("myTextList", "HMITextList")
    With objTextList
        .SelTextColor = RGB(255, 255, 0)
    End With
End Sub
