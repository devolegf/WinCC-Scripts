Sub PieSegmentConfiguration()
'VBA727
    Dim PieSegment As HMIPieSegment
    Set PieSegment = ActiveDocument.HMIObjects.AddHMIObject("PieSegment1", "HMIPieSegment")
    With PieSegment
        .StartAngle = 40
        .EndAngle = 180
    End With
End Sub
