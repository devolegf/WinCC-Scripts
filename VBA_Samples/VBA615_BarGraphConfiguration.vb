Sub BarGraphConfiguration()
'VBA615
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .LongStrokesTextEach = 3
    End With
End Sub
