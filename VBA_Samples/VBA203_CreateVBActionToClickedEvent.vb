Sub CreateVBActionToClickedEvent()
'VBA203
    Dim objButton As HMIButton
    Dim objCircle As HMICircle
    Dim objVBScript As HMIScriptInfo
    Dim strVBCode As String
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
        .Text = "Increase Radius"
    End With
    'define event and assign sourcecode to it:
    Set objVBScript = objButton.Events(1).Actions.AddAction(hmiActionCreationTypeVBScript)
    strVBCode = "Dim myCircle" & vbCrLf & "Set myCircle = "
    strVBCode = strVBCode & "HMIRuntime.ActiveScreen.ScreenItems(""Circle_VB"")"
    strVBCode = strVBCode & vbCrLf & "myCircle.Radius = myCircle.Radius + 5"
    With objVBScript
        .SourceCode = strVBCode
    End With
End Sub
