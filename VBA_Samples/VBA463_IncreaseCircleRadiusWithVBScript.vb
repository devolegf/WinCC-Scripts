Sub IncreaseCircleRadiusWithVBScript()
'VBA463
    Dim objButton As HMIButton
    Dim objCircleA As HMICircle
    Dim objEvent As HMIEvent
    Dim objVBScript As HMIScriptInfo
    Dim strCode As String
    strCode = "Dim objCircle" & vbCrLf & "Set objCircle = "
    strCode = strCode & "hmiRuntime.ActiveScreen.ScreenItems(""CircleVB"")"
    strCode = strCode & vbCrLf & "objCircle.Radius = objCircle.Radius + 5"
    Set objCircleA = ActiveDocument.HMIObjects.AddHMIObject("CircleVB", "HMICircle")
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    With objCircleA
        .Top = 100
        .Left = 100
    End With
    With objButton
        .Top = 10
        .Left = 10
        .Width = 200
        .Text = "Increase Radius"
    End With
    'On every mouseclick the radius will be increased:
    Set objEvent = objButton.Events(1)
    Set objVBScript = objButton.Events(1).Actions.AddAction(hmiActionCreationTypeVBScript)
    objVBScript.SourceCode = strCode
    Select Case objVBScript.Compiled
        Case True
            MsgBox "Compilation OK!"
        Case False
            MsgBox "Errors by compilation!"
    End Select
End Sub
