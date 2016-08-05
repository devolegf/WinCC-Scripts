Sub GroupDisplayConfiguration()
'VBA622
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCGUBackColorOn = RGB(255, 255, 255)
    End With
End Sub
