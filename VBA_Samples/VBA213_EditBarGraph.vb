Sub EditBarGraph()
'VBA213
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects("Bar1")
    objBarGraph.BorderColor = RGB(255, 0, 0)
End Sub
