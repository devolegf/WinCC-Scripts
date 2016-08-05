Sub IOFieldConfiguration()
'VBA528
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .HiddenInput = True
    End With
End Sub
