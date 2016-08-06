VBS80
Dim objScreen
Dim lngIndex
Dim lngAnswer
Dim strName
lngIndex = 1
Set objScreen = HMIRuntime.Screens("NewPDL1")
For lngIndex = 1 To objScreen.ScreenItems.Count
    strName = objScreen.ScreenItems.Item(lngIndex).ObjectName
    lngAnswer = MsgBox("Name of object " & lngIndex & ": " & strName, vbOKCancel)
    If vbCancel = lngAnswer Then Exit For
Next