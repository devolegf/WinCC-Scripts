Sub IOFieldConfiguration()
'VBA657
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .OperationReport = True
        .OperationMessage = True
    End With
End Sub
