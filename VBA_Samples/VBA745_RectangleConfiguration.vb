Sub RectangleConfiguration()
'VBA745
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .ToolTipText = "This is a rectangle"
    End With
End Sub
