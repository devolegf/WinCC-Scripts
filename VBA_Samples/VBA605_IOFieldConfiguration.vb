Sub IOFieldConfiguration()
'VBA605
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .DataFormat = 1
        .LimitMax = 100
    End With
End Sub
