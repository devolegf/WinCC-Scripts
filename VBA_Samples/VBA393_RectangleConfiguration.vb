Sub RectangleConfiguration()
'VBA393
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .BackColor = RGB(255, 255, 0)
    End With
End Sub
