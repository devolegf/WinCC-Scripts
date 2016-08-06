VBS75
Dim objScreen
Dim objCircle
Dim lngIndex
Dim strName
lngIndex = 1
Set objScreen = HMIRuntime.Screens("NewPDL1")
For lngIndex = 1 To objScreen.ScreenItems.Count
'
    'Searching all circles
    strName = objScreen.ScreenItems.Item(lngIndex).ObjectName
    If "Circle" = Left(strName, 6) Then
'
        'to halve the height of the circles
        Set objCircle = objScreen.ScreenItems(strName)
        objCircle.Height = objCircle.Height / 2
    End If
Next