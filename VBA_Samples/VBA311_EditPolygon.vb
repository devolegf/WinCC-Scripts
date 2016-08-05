Sub EditPolygon()
    Dim objPolygon As HMIPolygon
    Set objPolygon = ActiveDocument.HMIObjects("Polygon")
    objPolygon.BorderColor = RGB (255, 0, 0)
End Sub