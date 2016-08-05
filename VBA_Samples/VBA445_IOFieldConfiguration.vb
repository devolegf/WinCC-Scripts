Sub IOFieldConfiguration()
'VBA445
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .ClearOnError = True
    End With
End Sub
