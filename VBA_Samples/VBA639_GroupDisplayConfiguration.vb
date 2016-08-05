Sub GroupDisplayConfiguration()
'VBA639
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MessageClass = 0
        .MCText = "Alarm High"
    End With
End Sub
