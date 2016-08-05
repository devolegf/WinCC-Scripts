Sub RectangleConfiguration()
'VBA420
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .BorderFlashColorOff = RGB(0, 0, 0)
    End With
End Sub
