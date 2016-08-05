Sub AddRoundRectangle()
'VBA326
    Dim objRoundRectangle As HMIRoundRectangle
    Set objRoundRectangle = ActiveDocument.HMIObjects.AddHMIObject("Roundrectangle1", "HMIRoundRectangle")
End Sub
