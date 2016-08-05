Sub RectangleConfiguration()
'VBA399
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .BackFlashColorOn = RGB(0, 0, 255)
    End With
End Sub
