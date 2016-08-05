Sub ShowConnectorInfo_Menu()
'VBA297
    Dim objMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
    Dim strDocName As String
    strDocName = Application.ApplicationDataPath & ActiveDocument.Name
    Set objMenu = Documents(strDocName).CustomMenus.InsertMenu(1, "ConnectorMenu", "Connector_Info")
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(1, "ShowConnectInfo", "Info Connector")
End Sub
 
Sub ShowConnectorInfo()
    Dim objConnector As HMIObjConnection
    Dim iStart As Integer
    Dim iEnd As Integer
    Dim strStart As String
    Dim strEnd As String
    Dim strObjStart As String
    Dim strObjEnd As String
    Set objConnector = ActiveDocument.HMIObjects("Connector1")
    iStart = objConnector.BottomConnectedConnectionPointIndex
    iEnd = objConnector.TopConnectedConnectionPointIndex
    strObjStart = objConnector.BottomConnectedObjectName
    strObjEnd = objConnector.TopConnectedObjectName
    Select Case iStart
        Case 0
            strStart = "top"
        Case 1
            strStart = "right"
        Case 2
            strStart = "bottom"
        Case 3
            strStart = "left"
    End Select
    Select Case iEnd
        Case 0
            strEnd = "top"
        Case 1
            strEnd = "right"
        Case 2
            strEnd = "bottom"
        Case 3
            strEnd = "left"
    End Select
    MsgBox "The selected connector links the objects " & vbCrLf & "'" & strObjStart & "' and '" & strObjEnd & "'" & vbCrLf & "Connected points: " & vbCrLf & strObjStart & ": " & strStart & vbCrLf & strObjEnd & ": " & strEnd
End Sub

Private Sub Document_MenuItemClicked(ByVal MenuItem As IHMIMenuItem)
    Select Case MenuItem.Key
        Case "ShowConnectInfo"
            Call ShowConnectorInfo
    End Select
End Sub
