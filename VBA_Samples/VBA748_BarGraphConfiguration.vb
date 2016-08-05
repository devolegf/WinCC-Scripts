Sub BarGraphConfiguration()
'VBA748
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .trend = True
        .TrendColor = RGB(255, 0, 0)
    End With
End Sub
