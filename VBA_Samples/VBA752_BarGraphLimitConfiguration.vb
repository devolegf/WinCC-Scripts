Sub BarGraphLimitConfiguration()
'VBA752
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis = absolute
        .TypeAlarmLow = False
        'Activate monitoring
        .CheckAlarmLow = True
        'Set barcolor = "yellow"
        .ColorAlarmLow = RGB(255, 255, 0)
        'Set lower limit = "10"
        .AlarmLow = 10
    End With
End Sub
