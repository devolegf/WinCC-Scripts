Sub BarGraphLimitConfiguration()
'VBA756
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis = absolute
        .TypeLimitLow5 = False
        'Activate monitoring
        .CheckLimitLow5 = True
        'Set barcolor = "white"
        .ColorLimitLow5 = RGB(255, 255, 255)
        'Set lower limit = "0"
        .LimitLow5 = 0
    End With
End Sub
