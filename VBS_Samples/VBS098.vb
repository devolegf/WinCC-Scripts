VBS98
Dim objScreen
MsgBox HMIRuntime.ActiveScreen.ObjectName    'Output of active screen
Set objScreen = HMIRuntime.Screens("ScreenWindow1")
objScreen.Activate    'Activate "ScreenWindow1"
MsgBox HMIRuntime.ActiveScreen.ObjectName    'New output of active screen