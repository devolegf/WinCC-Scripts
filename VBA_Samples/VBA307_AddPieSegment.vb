Sub AddPieSegment()
'VBA307
    Dim objPieSegment As HMIPieSegment
    Set objPieSegment = ActiveDocument.HMIObjects.AddHMIObject("PieSegment1", "HMIPieSegment")
End Sub
