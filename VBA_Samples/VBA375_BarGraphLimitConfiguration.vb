Sub BarGraphLimitConfiguration()
'VBA375
    Dim objBarGraph As HMIBarGraph
'
    'Add new BarGraph to active document:
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolut
        .TypeAlarmHigh = False
        'Activate monitoring
        .CheckAlarmHigh = True
        'Set barcolor to "yellow"
        .ColorAlarmHigh = RGB(255, 255, 0)
        'set upper limit to "50"
        .AlarmHigh = 50
    End With
End Sub
