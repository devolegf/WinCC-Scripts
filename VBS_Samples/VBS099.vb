VBS99
Dim objScreen
Dim objScrItem
Dim lngIndex
Dim lngAnswer
Dim strName
lngIndex = 1
Set objScreen = HMIRuntime.Screens("NewPDL1")
For lngIndex = 1 To objScreen.ScreenItems.Count
'
    'The objects will be indicate by Item()
    strName = objScreen.ScreenItems.Item(lngIndex).ObjectName
    Set objScrItem = objScreen.ScreenItems(strName)
    lngAnswer = MsgBox(objScrItem.ObjectName, vbOKCancel)
    If vbCancel = lngAnswer Then Exit For
Next