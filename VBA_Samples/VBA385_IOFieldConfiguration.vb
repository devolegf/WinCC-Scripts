Sub IOFieldConfiguration()
'VBA385
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .AssumeOnExit = True
    End With
End Sub
