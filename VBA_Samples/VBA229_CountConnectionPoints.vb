Sub CountConnectionPoints()
'VBA229
    Dim objRectangle As HMIRectangle
    Dim objConnPoints As HMIConnectionPoints
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    Set objConnPoints = ActiveDocument.HMIObjects("Rectangle1").ConnectionPoints
    MsgBox "Rectangle1 has " & objConnPoints.Count & " connectionpoints."
End Sub
