Sub EditPieSegment()
'VBA308
    Dim objPieSegment As HMIPieSegment
    Set objPieSegment = ActiveDocument.HMIObjects("PieSegment1")
    objPieSegment.BorderColor = RGB(255, 0, 0)
End Sub
