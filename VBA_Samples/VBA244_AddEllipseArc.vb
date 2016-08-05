Sub AddEllipseArc()
'VBA244
    Dim objEllipseArc As HMIEllipseArc
    Set objEllipseArc = ActiveDocument.HMIObjects.AddHMIObject("EllipseArc", "HMIEllipseArc")
End Sub
