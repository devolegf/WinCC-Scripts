Sub TextListConfiguration()
'VBA545
    Dim objTextList As HMITextList
    Set objTextList = ActiveDocument.HMIObjects.AddHMIObject("myTextList", "HMITextList")
    With objTextList
        .ItemBorderWidth = 4
    End With
End Sub
