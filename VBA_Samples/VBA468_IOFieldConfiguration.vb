Sub IOFieldConfiguration()
'VBA468
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    Application.ActiveDocument.CursorMode = True
    With objIOField
        .CursorControl = True
    End With
End Sub
