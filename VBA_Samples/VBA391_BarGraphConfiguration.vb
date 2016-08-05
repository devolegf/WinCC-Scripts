Sub BarGraphConfiguration()
'VBA391
    Dim objBar As HMIBarGraph
    Set objBar = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBar
        .AxisSection = 1
    End With
End Sub
