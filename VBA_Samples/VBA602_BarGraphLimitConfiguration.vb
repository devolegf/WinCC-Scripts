Sub BarGraphLimitConfiguration()
'VBA602
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeLimitHigh5 = False
        'Activate monitoring
        .CheckLimitHigh5 = True
        'Set barcolor to "black"
        .ColorLimitHigh5 = RGB(0, 0, 0)
        'Set upper limit to "80"
        .LimitHigh4 = 80
    End With
End Sub
