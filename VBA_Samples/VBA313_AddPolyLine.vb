Sub AddPolyLine()
'VBA313
    Dim objPolyLine As HMIPolyLine
    Set objPolyLine = ActiveDocument.HMIObjects.AddHMIObject("PolyLine1", "HMIPolyLine")
End Sub
