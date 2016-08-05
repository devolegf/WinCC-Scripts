Sub GroupDisplayConfiguration()
'VBA430
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .Button4Width = 50
    End With
End Sub
