Sub EditRectangle()
'VBA321
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects("Rectangle1")
    objRectangle.BorderColor = RGB(255, 0, 0)
End Sub
