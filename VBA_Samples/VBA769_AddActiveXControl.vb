Sub AddActiveXControl()
'VBA769
    Dim objActiveXControl As HMIActiveXControl
    Set objActiveXControl = ActiveDocument.HMIObjects.AddActiveXControl("WinCC_Gauge2", "XGAUGE.XGaugeCtrl.1")
'
    'Move ActiveX-Control:
    objActiveXControl.Top = 40
    objActiveXControl.Left = 60
'
    'Modify individual properties:
    objActiveXControl.Properties("BackColor").value = RGB(255, 0, 0)
End Sub
