Sub Edit3DBarGraph()
'VBA201
    Dim obj3DBarGraph As HMI3DBarGraph
    Set obj3DBarGraph = ActiveDocument.HMIObjects("3DBar")
    obj3DBarGraph.BorderColor = RGB(255, 0, 0)
End Sub
