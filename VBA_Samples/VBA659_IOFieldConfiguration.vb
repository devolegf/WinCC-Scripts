Sub IOFieldConfiguration()
'VBA659
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .DataFormat = 1
        .OutputFormat = "99,999"
    End With
End Sub
