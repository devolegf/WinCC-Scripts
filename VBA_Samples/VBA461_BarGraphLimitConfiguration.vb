Sub BarGraphLimitConfiguration()
'VBA461
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeWarningLow = False
        'Activate monitoring
        .CheckWarningLow = True
        'Set barcolor to "magenta"
        .ColorWarningLow = RGB(255, 0, 255)
        'Set lower limit to "12"
        .WarningLow = 12
    End With
End Sub
