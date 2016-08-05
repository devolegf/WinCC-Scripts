Sub BarGraphConfiguration()
'VBA799
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .Scaling = True
        .ScalingType = 2
        .ZeroPoint = 50
        .ZeroPointValue = 0
    End With
End Sub
