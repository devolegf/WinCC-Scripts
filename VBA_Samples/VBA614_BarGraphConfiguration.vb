Sub BarGraphConfiguration()
'VBA614
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .LongStrokesSize = 10
    End With
End Sub
