Sub CreateVBActionToClickedEvent()
'VBA67
    Dim objButton As HMIButton
    Dim objCircle As HMICircle
    Dim objEvent As HMIEvent
    Dim objVBScript As HMIScriptInfo
    Dim strCode As String
    strCode = "Dim myCircle" & vbCrLf & "Set myCircle = "
    strCode = strCode & "HMIRuntime.ActiveScreen.ScreenItems(""Circle_VB"")"
    strCode = strCode & vbCrLf & "myCircle.Radius = myCircle.Radius + 5"
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_VB", "HMICircle")
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    With objCircle
        .Top = 100
        .Left = 100
        .BackColor = RGB(255, 0, 0)
    End With
    With objButton
        .Top = 10
        .Left = 10
        .Width = 120
        .Text = "Increase Radius"
    End With
    'Define event and assign sourcecode:
    Set objVBScript = objButton.Events(1).Actions.AddAction(hmiActionCreationTypeVBScript)
    With objVBScript
        .SourceCode = strCode
    End With
End Sub
