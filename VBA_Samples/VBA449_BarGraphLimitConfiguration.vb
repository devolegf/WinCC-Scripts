Sub BarGraphLimitConfiguration()
'VBA449
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeAlarmHigh = False
        'Activate monitoring
        .CheckAlarmHigh = True
        'Set barcolor to "red"
        .ColorAlarmHigh = RGB(255, 0, 0)
        'Set upper limit to "50"
        .AlarmHigh = 50
    End With
End Sub
