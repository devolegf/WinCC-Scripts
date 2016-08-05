Sub BarGraphLimitConfiguration()
'VBA754
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis = absolute
        .TypeLimitHigh5 = False
        'Activate monitoring
        .CheckLimitHigh5 = True
        'Set barcolor = "black"
        .ColorLimitHigh5 = RGB(0, 0, 0)
        'Set upper limit = "70"
        .LimitHigh5 = 70
    End With
End Sub
