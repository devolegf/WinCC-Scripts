Sub BarGraphConfiguration()
'VBA394
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .BackColor2 = RGB(255, 255, 0)
    End With
End Sub
