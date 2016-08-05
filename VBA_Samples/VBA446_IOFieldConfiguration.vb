Sub IOFieldConfiguration()
'VBA446
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .ClearOnNew = True
    End With
End Sub
