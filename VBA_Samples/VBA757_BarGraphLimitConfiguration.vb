Sub BarGraphLimitConfiguration()
'VBA757
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis = absolute
        .TypeToleranceHigh = False
        'Activate monitoring
        .CheckToleranceHigh = True
        'Set barcolor = "yellow"
        .ColorToleranceHigh = RGB(255, 255, 0)
        'Set upper limit = "40"
        .ToleranceHigh = 40
    End With
End Sub
