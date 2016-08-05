Sub HMI3DBarGraphConfiguration()
'VBA401
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .BarDepth = 40
    End With
End Sub
