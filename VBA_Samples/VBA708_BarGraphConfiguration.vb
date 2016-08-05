Sub BarGraphConfiguration()
'VBA708
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .ScalingType = 0
        .Scaling = True
    End With
End Sub
