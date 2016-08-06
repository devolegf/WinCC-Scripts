VBS141
Dim objRectangle
Set objRectangle = HMIRuntime.Screens("BaseScreen.ScreenWindow1:Screen1.ScreenWindow1:Screen2").ScreenItems("Rectangle1")
objRectangle.BackColor = RGB(255,0,0)