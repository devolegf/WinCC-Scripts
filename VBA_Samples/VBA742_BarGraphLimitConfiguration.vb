Sub BarGraphLimitConfiguration()
'VBA742
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis = absolute
        .TypeToleranceHigh = False
        'Activate monitoring
        .CheckToleranceHigh = True
        'Set barcolor = "yellow"
        .ColorToleranceHigh = RGB(255, 255, 0)
        'Set upper limit to "40"
        .ToleranceHigh = 40
    End With
End Sub
