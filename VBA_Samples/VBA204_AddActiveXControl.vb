Sub AddActiveXControl()
'VBA204
    Dim objActiveXControl As HMIActiveXControl
    Set objActiveXControl = ActiveDocument.HMIObjects.AddActiveXControl("WinCC_Gauge", "XGAUGE.XGaugeCtrl.1")
    With ActiveDocument
        .HMIObjects("WinCC_Gauge").Top = 40
        .HMIObjects("WinCC_Gauge").Left = 40
    End With
End Sub
