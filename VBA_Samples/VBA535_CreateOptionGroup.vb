Sub CreateOptionGroup()
'VBA535
    Dim objRadioBox As HMIOptionGroup
    Dim iIndex As Integer
    Set objRadioBox = ActiveDocument.HMIObjects.AddHMIObject("RadioBox_1", "HMIOptionGroup")
    With objRadioBox
        .Height = 100
        .Width = 180
        .BoxCount = 4
        For iIndex = 1 To .BoxCount
            .index = iIndex
            .Text = "myCustomText" & .index
        Next iIndex
    End With
End Sub
