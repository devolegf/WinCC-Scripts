Sub RectangleConfiguration()
'VBA746
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .Left = 10
        .Top = 40
    End With
End Sub
