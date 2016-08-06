VBS79
Dim objScreen
Dim objScrItem
Dim lngIndex
Dim strName
lngIndex = 1
Set objScreen = HMIRuntime.Screens("NewPDL1")
For lngIndex = 1 To objScreen.ScreenItems.Count
    strName = objScreen.ScreenItems.Item(lngIndex).ObjectName
    Set objScrItem = objScreen.ScreenItems(strName)
    objScrItem.Left = objScrItem.Left - 5
Next