Sub EditDefaultPropertiesOfEllipseArc()
'VBA247
    Dim objEllipseArc As HMIEllipseArc
    Set objEllipseArc = Application.DefaultHMIObjects("HMIEllipseArc")
    objEllipseArc.BorderColor = RGB(255, 255, 0)
    'create new "EllipseArc"-object
    Set objEllipseArc = ActiveDocument.HMIObjects.AddHMIObject("EllipseArc2", "HMIEllipseArc")
End Sub
