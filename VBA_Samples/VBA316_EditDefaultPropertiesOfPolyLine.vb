Sub EditDefaultPropertiesOfPolyLine()
'VBA316
    Dim objPolyLine As HMIPolyLine
    Set objPolyLine = Application.DefaultHMIObjects("HMIPolyLine")
    objPolyLine.BorderColor = RGB(255, 255, 0)
End Sub
