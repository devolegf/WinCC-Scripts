VBS71
Dim objScreen
Dim objScrItem
Dim lngIndex
Dim strName
lngIndex = 1
Set objScreen = HMIRuntime.Screens("NewPDL1")
For lngIndex = 1 To objScreen.ScreenItems.Count
    strName = objScreen.ScreenItems.Item(lngIndex).ObjectName    'Read names of objects
    Set objScrItem = objScreen.ScreenItems(strName)
    objScrItem.Enabled=False    'Lock object
Next