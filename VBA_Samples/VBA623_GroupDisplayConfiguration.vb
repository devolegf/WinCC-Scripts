Sub GroupDisplayConfiguration()
'VBA623
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCGUBackFlash = True
    End With
End Sub
