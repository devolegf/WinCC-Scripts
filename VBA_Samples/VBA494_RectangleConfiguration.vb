Sub RectangleConfiguration()
'VBA494
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .Filling = True
    End With
End Sub
