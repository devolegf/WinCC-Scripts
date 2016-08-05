Sub BarGraphConfiguration()
'VBA531
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .Hysteresis = True
    End With
End Sub
