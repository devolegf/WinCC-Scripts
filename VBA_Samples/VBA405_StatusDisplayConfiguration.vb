Sub StatusDisplayConfiguration()
'VBA405
    Dim objStatDisp As HMIStatusDisplay
    Set objStatDisp = ActiveDocument.HMIObjects.AddHMIObject("Statusdisplay1", "HMIStatusDisplay")
    With objStatDisp
        .BasePicTransColor = RGB(255, 255, 0)
        .BasePicUseTransColor = True
    End With
End Sub
