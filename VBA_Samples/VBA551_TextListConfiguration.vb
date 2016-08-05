Sub TextListConfiguration()
'VBA551
    Dim objTextList As HMITextList
    Set objTextList = ActiveDocument.HMIObjects.AddHMIObject("myTextList", "HMITextList")
    With objTextList
        .LanguageSwitch = True
    End With
End Sub
