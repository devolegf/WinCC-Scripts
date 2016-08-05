Sub BarGraphConfiguration()
'VBA613
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .LongStrokesOnly = True
    End With
End Sub
