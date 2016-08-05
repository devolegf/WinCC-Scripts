Sub IOFieldConfiguration()
'VBA379
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .AlignmentLeft = 1
        .AlignmentTop = 1
    End With
End Sub
