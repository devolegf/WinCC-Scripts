Sub CreateOptionGroup()
'VBA425
    Dim objRadioBox As HMIOptionGroup
    Dim iCounter As Integer
    Set objRadioBox = ActiveDocument.HMIObjects.AddHMIObject("RadioBox_1", "HMIOptionGroup")
    iCounter = 1
    With objRadioBox
        .Height = 100
        .Width = 180
        .BoxCount = 4
        .BoxAlignment = True
        For iCounter = 1 To .BoxCount
            .index = iCounter
            .Text = "CustomText" & .index
        Next iCounter
    End With
End Sub
