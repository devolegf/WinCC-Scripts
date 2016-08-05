Sub IOFieldConfiguration()
'VBA484
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .EditAtOnce = True
    End With
End Sub
