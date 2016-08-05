Sub GroupDisplayConfiguration()
'VBA427
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .Button1Width = 50
    End With
End Sub
