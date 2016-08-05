Sub GroupDisplayConfiguration()
'VBA428
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .Button2Width = 50
    End With
End Sub
