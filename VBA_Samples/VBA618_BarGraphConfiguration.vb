Sub BarGraphConfiguration()
'VBA618
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .Max = 10
    End With
End Sub
