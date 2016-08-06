VBS82
Dim objScrItem
Set objScrItem = HMIRuntime.Screens(1).ScreenItems(1)
MsgBox "Name of BaseScreen: " & objScrItem.Parent.ObjectName