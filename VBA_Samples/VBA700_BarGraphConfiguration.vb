Sub BarGraphConfiguration()
'VBA700
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .RightComma = 4
    End With
End Sub
