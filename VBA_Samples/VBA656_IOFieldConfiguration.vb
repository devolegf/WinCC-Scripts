Sub IOFieldConfiguration()
'VBA656
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .OperationReport = True
        .OperationMessage = True
    End With
End Sub
