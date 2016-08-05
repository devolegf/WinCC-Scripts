Sub BarGraphLimitConfiguration()
'VBA441
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeToleranceHigh = False
        'Activate monitoring
        .CheckToleranceHigh = True
        'Set barcolor to "yellow"
        .ColorToleranceHigh = RGB(255, 255, 0)
        'Set upper limit to "45"
        .ToleranceHigh = 45
    End With
End Sub
