Sub BarGraphConfiguration()
'VBA706
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .Scaling = True
        .ScaleTicks = 10
    End With
End Sub
