Sub AddActiveXControl()
'VBA717
    Dim objActiveXControl As HMIActiveXControl
    Set objActiveXControl = ActiveDocument.HMIObjects.AddActiveXControl("WinCC_Gauge", "XGAUGE.XGaugeCtrl.1")
    With objActiveXControl
        .Top = 40
        .Left = 60
        MsgBox .Properties("ServerName").value
    End With
End Sub
