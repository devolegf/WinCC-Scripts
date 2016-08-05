Sub BarGraphLimitConfiguration()
'VBA439
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeLimitLow4 = False
        'Activate monitoring
        .CheckLimitLow4 = True
        'Set barcolor to "green"
        .ColorLimitLow4 = RGB(0, 255, 0)
        'set lower limit to "5"
        .LimitLow4 = 5
    End With
End Sub
