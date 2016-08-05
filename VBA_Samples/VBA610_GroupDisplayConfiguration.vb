Sub GroupDisplayConfiguration()
'VBA610
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .LockStatus = True
        .LockText = "gesperrt"
    End With
End Sub
