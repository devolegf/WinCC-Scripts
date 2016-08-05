Sub BarGraphConfiguration()
'VBA599
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .LeftComma = 4
    End With
End Sub
