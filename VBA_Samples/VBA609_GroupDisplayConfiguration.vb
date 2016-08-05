Sub GroupDisplayConfiguration()
'VBA609
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .LockStatus = True
        .LockBackColor = RGB(255, 0, 0)
    End With
End Sub
