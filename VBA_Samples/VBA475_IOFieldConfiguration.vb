Sub IOFieldConfiguration()
'VBA475
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .DataFormat = 1
    End With
End Sub
