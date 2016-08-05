Sub BarGraphLimitConfiguration()
'VBA751
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis = absolute
        .TypeAlarmHigh = False
        'Activate monitoring
        .CheckAlarmHigh = True
        'Set barcolor = "yellow"
        .ColorAlarmHigh = RGB(255, 255, 0)
        'Set upper limit = "50"
        .AlarmHigh = 50
    End With
End Sub
