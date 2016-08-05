Sub EditCiruclarArc()
'VBA227
    Dim objCiruclarArc As HMICircularArc
    Set objCiruclarArc = ActiveDocument.HMIObjects("CircularArc")
    objCiruclarArc.BorderColor = RGB(255, 0, 0)
End Sub
