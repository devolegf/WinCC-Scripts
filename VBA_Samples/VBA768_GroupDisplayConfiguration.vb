Sub GroupDisplayConfiguration()
'VBA768
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .UserValue1 = 0
        .UserValue2 = 25
        .UserValue3 = 50
        .UserValue4 = 75
    End With
End Sub
