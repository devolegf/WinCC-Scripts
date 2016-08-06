VBS96
Dim objScreen
Dim cmdButton
Dim lngIndex
Dim strName
lngIndex = 1
Set objScreen = HMIRuntime.Screens("NewPDL1")
For lngIndex = 1 To objScreen.ScreenItems.Count
'
    'Get all "Buttons"
    strName = objScreen.ScreenItems(lngIndex).ObjectName
    If "Button" = Left(strName, 6) Then
        Set cmdButton = objScreen.ScreenItems(strName)
        cmdButton.Width = cmdButton.Width * 2
    End If
Next