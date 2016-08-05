Sub GroupDisplayConfiguration()
'VBA621
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCGUBackColorOff = RGB(255, 0, 0)
    End With
End Sub
