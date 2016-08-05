Sub BarGraphConfiguration()
'VBA377
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .Alignment = True
        .Scaling = True
    End With
End Sub
