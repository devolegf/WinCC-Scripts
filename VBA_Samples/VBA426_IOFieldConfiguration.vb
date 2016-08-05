Sub IOFieldConfiguration()
'VBA426
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .BoxType = 1
    End With
End Sub
