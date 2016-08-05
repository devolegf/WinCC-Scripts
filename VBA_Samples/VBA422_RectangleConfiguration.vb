Sub RectangleConfiguration()
'VBA422
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .BorderStyle = 1
    End With
End Sub
