Sub EditEllipseArc()
'VBA245
    Dim objEllipseArc As HMIEllipseArc
    Set objEllipseArc = ActiveDocument.HMIObjects("EllipseArc")
    objEllipseArc.BorderColor = RGB(255, 0, 0)
End Sub
