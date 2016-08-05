Sub StatusDisplayConfiguration()
'VBA404
    Dim objStatDisp As HMIStatusDisplay
    Set objStatDisp = ActiveDocument.HMIObjects.AddHMIObject("Statusdisplay1", "HMIStatusDisplay")
    With objStatDisp
        .BasePicReferenced = True
    End With
End Sub
