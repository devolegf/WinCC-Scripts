Sub BarGraphConfiguration()
'VBA395
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .BackColor3 = RGB(0, 0, 255)
    End With
End Sub
