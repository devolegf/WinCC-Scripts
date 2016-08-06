VBS78
Dim objScreen
Dim objScrItem
Dim lngAnswer
Dim lngIndex
Dim strName
lngIndex = 1
Set objScreen = HMIRuntime.Screens("NewPDL1")
For lngIndex = 1 To objScreen.ScreenItems.Count
    strName = objScreen.ScreenItems.Item(lngIndex).ObjectName
    Set objScrItem = objScreen.ScreenItems(strName)
    lngAnswer = MsgBox(strName & " is in layer  " & objScrItem.Layer,vbOKCancel)
    If vbCancel = lngAnswer Then Exit For
Next