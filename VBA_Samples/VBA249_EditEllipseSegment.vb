Sub EditEllipseSegment()
'VBA249
    Dim objEllipseSegment As HMIEllipseSegment
    Set objEllipseSegment = ActiveDocument.HMIObjects("EllipseSegment")
    objEllipseSegment.BorderColor = RGB(255, 0, 0)
End Sub
