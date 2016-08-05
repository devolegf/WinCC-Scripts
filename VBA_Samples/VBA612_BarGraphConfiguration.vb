Sub BarGraphConfiguration()
'VBA612
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .LongStrokesBold = False
    End With
End Sub
