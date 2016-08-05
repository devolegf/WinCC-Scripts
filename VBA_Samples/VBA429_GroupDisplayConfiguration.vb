Sub GroupDisplayConfiguration()
'VBA429
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .Button3Width = 50
    End With
End Sub
