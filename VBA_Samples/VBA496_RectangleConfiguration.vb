Sub RectangleConfiguration()
'VBA496
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .FillStyle = 196643
    End With
End Sub
