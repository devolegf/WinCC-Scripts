Sub IOFieldConfig()
'VBA732
    Dim objIOField1 As HMIIOField
    Dim objIOField2 As HMIIOField
    Set objIOField1 = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    Set objIOField2 = ActiveDocument.HMIObjects.AddHMIObject("IOField2", "HMIIOField")
    With objIOField1
        .Top = 10
        .Left = 10
        .TabOrderSwitch = 1
    End With
    With objIOField2
        .Top = 100
        .Left = 10
        .TabOrderSwitch = 2
    End With
End Sub
