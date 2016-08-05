Sub BarGraphLimitConfiguration()
'VBA601
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeLimitHigh4 = False
        'Activate monitoring
        .CheckLimitHigh4 = True
        'Set barcolor to "red"
        .ColorLimitHigh4 = RGB(255, 0, 0)
        'Set upper limit to "70"
        .LimitHigh4 = 70
    End With
End Sub
