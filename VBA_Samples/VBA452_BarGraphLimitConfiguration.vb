Sub BarGraphLimitConfiguration()
'VBA452
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .ColorChangeType = False
    End With
End Sub
