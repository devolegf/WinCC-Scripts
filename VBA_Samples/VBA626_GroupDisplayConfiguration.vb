Sub GroupDisplayConfiguration()
'VBA626
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCGUTextFlash = True
    End With
End Sub
