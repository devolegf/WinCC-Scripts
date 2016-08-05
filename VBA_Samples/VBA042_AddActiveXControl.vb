Sub AddActiveXControl()
'VBA42
    Dim objActiveXControl As HMIActiveXControl
    Set objActiveXControl = ActiveDocument.HMIObjects.AddActiveXControl("WinCC_Gauge2", "XGAUGE.XGaugeCtrl.1")
'
    'move ActiveX-control:
    objActiveXControl.Top = 40
    objActiveXControl.Left = 60
'
    'Change individual property:
    objActiveXControl.Properties("BackColor").value = RGB(255, 0, 0)
End Sub
