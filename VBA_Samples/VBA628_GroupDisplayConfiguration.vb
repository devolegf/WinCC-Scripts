Sub GroupDisplayConfiguration()
'VBA628
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCKOBackColorOn = RGB(255, 255, 255)
    End With
End Sub
