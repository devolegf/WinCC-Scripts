Sub ShowCustomMenusOfDocument()
'VBA290
    Dim colMenus As HMIMenus
    Dim objMenu As HMIMenu
    Dim strMenuList As String
    Set colMenus = ActiveDocument.CustomMenus
    For Each objMenu In colMenus
        strMenuList = strMenuList & objMenu.Label & vbCrLf
    Next objMenu
    MsgBox strMenuList
End Sub
