Sub BarGraphLimitConfiguration()
'VBA794
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis = absolute
        .TypeWarningHigh = False
        'Activate monitoring
        .CheckWarningHigh = True
        'Set barcolor = "red"
        .ColorWarningHigh = RGB(255, 0, 0)
        'Set upper limit = "75"
        .WarningHigh = 75
    End With
End Sub
