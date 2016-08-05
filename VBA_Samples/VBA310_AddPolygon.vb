Sub AddPolygon()
'VBA310
    Dim objPolygon As HMIPolygon
    Set objPolygon = ActiveDocument.HMIObjects.AddHMIObject("Polygon", "HMIPolygon")
End Sub
