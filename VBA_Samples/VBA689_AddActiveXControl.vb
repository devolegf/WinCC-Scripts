Sub AddActiveXControl()
'VBA689
    Dim objActiveXControl As HMIActiveXControl
    Set objActiveXControl = ActiveDocument.HMIObjects.AddActiveXControl("WinCC_Gauge", "XGAUGE.XGaugeCtrl.1")
    With ActiveDocument
        .HMIObjects("WinCC_Gauge").Top = 40
        .HMIObjects("WinCC_Gauge").Left = 40
        MsgBox "ProgID of ActiveX-control: " & .HMIObjects("WinCC_Gauge").ProgID
    End With
End Sub
