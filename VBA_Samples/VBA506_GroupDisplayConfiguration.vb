Sub GroupDisplayConfiguration()
'VBA506
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .FlashRate = 1
    End With
End Sub
