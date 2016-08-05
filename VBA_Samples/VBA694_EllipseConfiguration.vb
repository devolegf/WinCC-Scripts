Sub EllipseConfiguration()
'VBA694
    Dim objEllipse As HMIEllipse
    Set objEllipse = ActiveDocument.HMIObjects.AddHMIObject("Ellipse1", "HMIEllipse")
    With objEllipse
        .RadiusHeight = 60
        .RadiusWidth = 40
    End With
End Sub
