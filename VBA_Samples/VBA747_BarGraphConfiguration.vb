Sub BarGraphConfiguration()
'VBA747
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .trend = True
    End With
End Sub
