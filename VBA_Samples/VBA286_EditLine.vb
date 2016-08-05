Sub EditLine()
'VBA286
    Dim objLine As HMILine
    Set objLine = ActiveDocument.HMIObjects("Line1")
    objLine.BorderColor = RGB(255, 0, 0)
End Sub
