Sub BarGraphLimitConfiguration()
'VBA376
    Dim objBarGraph As HMIBarGraph
'
    'Add new bargraph to active document:
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolut
        .TypeAlarmLow = False
        'Activate monitoring
        .CheckAlarmLow = True
        'Set Barcolor to "yellow"
        .ColorAlarmLow = RGB(255, 255, 0)
        'set lower limit to "10"
        .AlarmLow = 10
    End With
End Sub
