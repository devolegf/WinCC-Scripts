Sub BarGraphConfiguration()
'VBA532
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .Hysteresis = True
        .HysteresisRange = 4
    End With
End Sub
