Sub BarGraphLimitConfiguration()
'VBA450
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeAlarmLow = False
        'Activate monitoring
        .CheckAlarmLow = True
        'Set barcolor to "red"
        .ColorAlarmLow = RGB(255, 0, 0)
        'Set lower limit to "10"
        .AlarmLow = 10
    End With
End Sub
