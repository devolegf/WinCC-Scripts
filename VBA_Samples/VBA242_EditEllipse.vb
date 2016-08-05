Sub EditEllipse()
'VBA242
    Dim objEllipse As HMIEllipse
    Set objEllipse = ActiveDocument.HMIObjects("Ellipse")
    objEllipse.BorderColor = RGB(255, 0, 0)
End Sub
