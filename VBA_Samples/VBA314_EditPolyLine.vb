Sub EditPolyLine()
'VBA314
    Dim objPolyLine As HMIPolyLine
    Set objPolyLine = ActiveDocument.HMIObjects("PolyLine1")
    objPolyLine.BorderColor = RGB(255, 0, 0)
End Sub
