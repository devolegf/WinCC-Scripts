Sub AddCircle()
'VBA223
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle", "HMICircle")
End Sub
