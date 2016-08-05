Sub AddEllipseSegment()
'VBA248
    Dim objEllipseSegment As HMIEllipseSegment
    Set objEllipseSegment = ActiveDocument.HMIObjects.AddHMIObject("EllipseSegment", "HMIEllipseSegment")
End Sub
