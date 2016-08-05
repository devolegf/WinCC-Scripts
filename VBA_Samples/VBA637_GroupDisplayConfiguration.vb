Sub GroupDisplayConfiguration()
'VBA637
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCKQTextColorOn = RGB(0, 0, 0)
    End With
End Sub
