Sub AddLine()
'VBA285
    Dim objLine As HMILine
    Set objLine = ActiveDocument.HMIObjects.AddHMIObject("Line1", "HMILine")
End Sub
