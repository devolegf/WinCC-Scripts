Sub GroupDisplayConfiguration()
'VBA634
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCKQBackColorOn = RGB(255, 255, 255)
    End With
End Sub
