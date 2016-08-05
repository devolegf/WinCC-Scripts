Sub BarGraphConfiguration()
'VBA497
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .FillStyle2 = 196643
    End With
End Sub
