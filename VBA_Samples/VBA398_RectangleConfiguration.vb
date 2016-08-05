Sub RectangleConfiguration()
'VBA398
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .BackFlashColorOff = RGB(255, 255, 0)
    End With
End Sub
