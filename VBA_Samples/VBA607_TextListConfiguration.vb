Sub TextListConfiguration()
'VBA607
    Dim objTextList As HMITextList
    Set objTextList = ActiveDocument.HMIObjects.AddHMIObject("myTextList", "HMITextList")
    With objTextList
        .ListType = 0
    End With
End Sub
