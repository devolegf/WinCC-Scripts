Sub BarGraphLimitConfiguration()
'VBA435
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeAlarmLow = False
        'Activate monitoring
        .CheckAlarmLow = True
        'Set barcolor to "yellow"
        .ColorAlarmLow = RGB (255, 255, 0)
        'Set lower limit to "10"
        .AlarmLow = 10
    End With
End Sub