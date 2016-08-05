Sub EditCircle()
'VBA224
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects("Circle")
    objCircle.BorderColor = RGB(255, 0, 0)
End Sub
