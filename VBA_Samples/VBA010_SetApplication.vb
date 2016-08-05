Option Explicit
'VBA10
'The next declaration has to be placed in the modul section
Dim WithEvents theApp As grafexe.Application


Private Sub SetApplication()
    'This procedure has to be executed (with "F5") first
    Set theApp = grafexe.Application
End Sub


Private Sub theApp_MenuItemClicked(ByVal MenuItem As IHMIMenuItem)
    Dim objClicked As HMIMenuItem
    Dim varMenuItemKey As Variant
    Set objClicked = MenuItem
'
    '"varMenuItemKey" contains the value of parameter "Key"
    'from clicked menu-item
    varMenuItemKey = objClicked.Key
    Select Case varMenuItemKey
        Case "mItem1_1"
            MsgBox "The first menuitem was clicked!"
    End Select
End Sub

Private Sub theApp_ToolbarItemClicked(ByVal ToolbarItem As IHMIToolbarItem)
    Dim objClicked As HMIToolbarItem
    Dim varToolbarItemKey As Variant
    Set objClicked = ToolbarItem
'
    '"varToolbarItemKey" contains the value of parameter "Key"
    'from clicked toolbar-item
    varToolbarItemKey = objClicked.Key
    Select Case varToolbarItemKey
        Case "tItem1_1"
            MsgBox "The first symbol-icon was clicked!"
    End Select
End Sub