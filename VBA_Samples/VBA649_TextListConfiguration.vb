Sub TextListConfiguration()
'VBA649
    Dim objTextList As HMITextList
    Set objTextList = ActiveDocument.HMIObjects.AddHMIObject("myTextList", "HMITextList")
    With objTextList
        .NumberLines = 3
    End With
End Sub
