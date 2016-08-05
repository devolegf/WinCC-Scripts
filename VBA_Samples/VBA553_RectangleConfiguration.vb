Sub RectangleConfiguration()
'VBA553
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .Layer = 4
    End With
End Sub
