Sub GroupDisplayConfiguration()
'VBA611
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .LockStatus = True
        .LockTextColor = RGB(0, 255, 255)
    End With
End Sub
