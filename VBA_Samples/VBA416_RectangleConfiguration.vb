Sub RectangleConfiguration()
'VBA416
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .BorderColor = RGB(0, 0, 255)
    End With
End Sub
