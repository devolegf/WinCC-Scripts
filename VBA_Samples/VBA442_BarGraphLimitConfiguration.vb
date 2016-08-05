Sub BarGraphLimitConfiguration()
'VBA442
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeToleranceLow = False
        'Activate monitoring
        .CheckToleranceLow = True
        'Set barcolor to "yellow"
        .ColorToleranceLow = RGB(255, 255, 0)
        'Set lower limit to "15"
        .ToleranceLow = 15
    End With
End Sub
