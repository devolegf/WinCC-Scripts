Sub IOFieldConfiguration()
'VBA660
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .OutputValue = "00"
    End With
End Sub
