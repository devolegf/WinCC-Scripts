Sub RectangleConfiguration()
'VBA421
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .BorderFlashColorOn = RGB(255, 0, 0)
    End With
End Sub
