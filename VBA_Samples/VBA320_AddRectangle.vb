Sub AddRectangle()
'VBA320
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
End Sub
