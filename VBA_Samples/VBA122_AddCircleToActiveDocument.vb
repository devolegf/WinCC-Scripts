Sub AddCircleToActiveDocument()
'VBA122
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("VBA_Circle", "HMICircle")
    objCircle.BackColor = RGB(255, 0, 0)
End Sub
