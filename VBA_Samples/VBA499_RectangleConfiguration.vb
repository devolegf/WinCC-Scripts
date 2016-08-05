Sub RectangleConfiguration()
'VBA499
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .FlashBorderColor = True
    End With
End Sub
