Sub BarGraphLimitConfiguration()
'VBA743
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis = absolute
        .TypeToleranceLow = False
        'Activate monitoring
        .CheckToleranceLow = True
        'Set barcolor = "red"
        .ColorToleranceLow = RGB(255, 0, 0)
        'Set lower limit to "40"
        .ToleranceLow = 40
    End With
End Sub
