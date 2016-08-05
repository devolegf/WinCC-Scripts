Sub BarGraphConfiguration()
'VBA490
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .Exponent = True
    End With
End Sub
