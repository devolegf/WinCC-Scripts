Sub RoundRectangleConfiguration()
'VBA702
    Dim objRoundRectangle As HMIRoundRectangle
    Set objRoundRectangle = ActiveDocument.HMIObjects.AddHMIObject("RoundRectangle1", "HMIRoundRectangle")
    With objRoundRectangle
        .RoundCornerHeight = 25
        .RoundCornerWidth = 50
    End With
End Sub
