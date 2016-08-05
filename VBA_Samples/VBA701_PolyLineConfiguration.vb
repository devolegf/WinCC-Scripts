Sub PolyLineConfiguration()
'VBA701
    Dim objPolyLine As HMIPolyLine
    Set objPolyLine = ActiveDocument.HMIObjects.AddHMIObject("PolyLine1", "HMIPolyLine")
    With objPolyLine
        .ReferenceRotationLeft = 50
        .ReferenceRotationTop = 50
        .RotationAngle = 45
    End With
End Sub
