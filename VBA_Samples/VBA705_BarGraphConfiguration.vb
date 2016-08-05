Sub BarGraphConfiguration()
'VBA705
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .Scaling = True
        .ScaleColor = RGB(255, 0, 0)
    End With
End Sub
