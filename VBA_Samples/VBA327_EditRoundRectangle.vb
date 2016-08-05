Sub EditRoundRectangle()
'VBA327
    Dim objRoundRectangle As HMIRoundRectangle
    Set objRoundRectangle = ActiveDocument.HMIObjects("Roundrectangle1")
    objRoundRectangle.BorderColor = RGB(255, 0, 0)
End Sub
