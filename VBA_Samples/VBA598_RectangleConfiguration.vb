Sub RectangleConfiguration()
'VBA598
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .Left = 40
    End With
End Sub
