Sub PolyLineConfiguration()
'VBA696
    Dim objPolyLine As HMIPolyLine
    Set objPolyLine = ActiveDocument.HMIObjects.AddHMIObject("PolyLine1", "HMIPolyLine")
    With objPolyLine
        .ReferenceRotationLeft = 50
        .ReferenceRotationTop = 50
    End With
End Sub
