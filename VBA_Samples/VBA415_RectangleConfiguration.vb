Sub RectangleConfiguration()
'VBA415
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .BorderBackColor = RGB(255, 255, 0)
    End With
End Sub
