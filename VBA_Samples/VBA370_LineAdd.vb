Sub LineAdd()
'VBA370
    Dim objLine As HMILine
    Set objLine = ActiveDocument.HMIObjects.AddHMIObject("myLine", "HMILine")
    With objLine
        .BorderColor = RGB(255, 0, 0)
        .index = hmiLineIndexTypeStartPoint
        .ActualPointLeft = 12
        .ActualPointTop = 34
        .index = hmiLineIndexTypeEndPoint
        .ActualPointLeft = 74
        .ActualPointTop = 64
    End With
End Sub
