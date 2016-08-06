VBS93
Dim objScreen
Dim objScrItem
Dim lngIndex
Dim lngAnswer
Dim strName
lngIndex = 1
Set objScreen = HMIRuntime.Screens("NewPDL1")
For lngIndex = 1 To objScreen.ScreenItems.Count
    strName = objScreen.ScreenItems(lngIndex).ObjectName
    Set objScrItem = objScreen.ScreenItems(strName)
    lngAnswer = MsgBox(objScrItem.Type, vbOKCancel)
    If vbCancel = lngAnswer Then Exit For
Next