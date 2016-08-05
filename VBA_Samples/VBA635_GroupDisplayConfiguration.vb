Sub GroupDisplayConfiguration()
'VBA635
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCKQBackFlash = True
    End With
End Sub
