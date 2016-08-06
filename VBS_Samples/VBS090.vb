VBS90
Dim objScreen
Dim objScrItem
Dim lngIndex
Dim strName
lngIndex = 1
Set objScreen = HMIRuntime.Screens("NewPDL1")
For lngIndex = 1 To objScreen.ScreenItems.Count
    strName = objScreen.ScreenItems(lngIndex).ObjectName
    Set objScrItem = objScreen.ScreenItems(strName)
    objScrItem.Top = objScrItem.Top - 5
Next