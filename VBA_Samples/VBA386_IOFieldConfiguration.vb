Sub IOFieldConfiguration()
'VBA386
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .AssumeOnFull = True
    End With
End Sub
