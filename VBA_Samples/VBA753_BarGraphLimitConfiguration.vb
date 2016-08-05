Sub BarGraphLimitConfiguration()
'VBA753
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis = absolute
        .TypeLimitHigh4 = False
        'Activate monitoring
        .CheckLimitHigh4 = True
        'Set barcolor = "red"
        .ColorLimitHigh4 = RGB(255, 0, 0)
        'Set upper limit = "70"
        .LimitHigh4 = 70
    End With
End Sub
