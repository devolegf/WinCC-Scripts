Sub HMI3DBarGraphConfiguration()
'VBA380
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        'Depth-angle a = 15 degrees
        .AngleAlpha = 15
        'Depth-angle b = 45 degrees
        .AngleBeta = 45
    End With
End Sub
