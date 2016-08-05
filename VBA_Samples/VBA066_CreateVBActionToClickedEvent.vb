Sub CreateVBActionToClickedEvent()
'VBA66
    Dim objButton As HMIButton
    Dim objCircle As HMICircle
    Dim objEvent As HMIEvent
    Dim objCScript As HMIScriptInfo
    Dim strCode As String
    strCode = "long lHeight;" & vbCrLf & "int check;" & vbCrLf
    strCode = strCode & "lHeight = GetHeight (""events.PDL"", ""myCircle"");"
    strCode = strCode & vbCrLf & "lHeight = lHeight+5;" & vbCrLf & "check = "
    strCode = strCode & "SetHeight(""events.PDL"", ""myCircle"",lHeight);"
    strCode = strCode & vbCrLf & "//Return-Type: Void"
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_C", "HMICircle")
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    With objCircle
        .Top = 100
        .Left = 100
        .BackColor = RGB(255, 0, 0)
    End With
    With objButton
        .Top = 10
        .Left = 10
        .Text = "Increase Radius"
    End With
    'Configure directconnection:
    Set objCScript = objButton.Events(1).Actions.AddAction(hmiActionCreationTypeCScript)
    With objCScript
'
        'Note: Replace "events.PDL" with your picturename
        .SourceCode = strCode
    End With
End Sub
