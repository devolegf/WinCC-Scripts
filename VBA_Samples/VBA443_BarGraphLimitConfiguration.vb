Sub BarGraphLimitConfiguration()
'VBA443
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeWarningHigh = False
        'Activate monitoring
        .CheckWarningHigh = True
        'Set barcolor to "red"
        .ColorWarningHigh = RGB(255, 0, 0)
        'Set upper limit to "75"
        .WarningHigh = 75
    End With
End Sub
