Sub AddEllipse()
'VBA241
    Dim objEllipse As HMIEllipse
    Set objEllipse = ActiveDocument.HMIObjects.AddHMIObject("Ellipse", "HMIEllipse")
End Sub
