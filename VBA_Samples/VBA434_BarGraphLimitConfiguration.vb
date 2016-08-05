Sub BarGraphLimitConfiguration()
'VBA434
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeAlarmHigh = False
        'Activate monitoring
        .CheckAlarmHigh = True
        'Set barcolor to "yellow"
        .ColorAlarmHigh = RGB(255, 255, 0)
        'Set upper limit to "50"
        .AlarmHigh = 50
    End With
End Sub
