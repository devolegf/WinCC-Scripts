Sub IOFieldConfiguration()
'VBA606
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .DataFormat = 1
        .LimitMin = 0
    End With
End Sub
