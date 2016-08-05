Sub RectangleConfiguration()
'VBA495
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .Filling = True
        .FillingIndex = 50
    End With
End Sub
