Sub PieSegmentConfiguration()
'VBA693
    Dim objPieSegment As HMIPieSegment
    Set objPieSegment = ActiveDocument.HMIObjects.AddHMIObject("PieSegment1", "HMIPieSegment")
    With objPieSegment
        .StartAngle = 40
        .EndAngle = 180
        .Radius = 80
    End With
End Sub
