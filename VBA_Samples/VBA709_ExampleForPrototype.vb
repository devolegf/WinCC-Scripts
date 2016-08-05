Sub ExampleForPrototype()
'VBA709
    Dim objButton As HMIButton
    Dim objCircleA As HMICircle
    Dim objEvent As HMIEvent
    Dim objVBScript As HMIScriptInfo
    Dim strScriptType As String
    Set objCircleA = ActiveDocument.HMIObjects.AddHMIObject("CircleA", "HMICircle")
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
    'On every mouseclick the radius have to increase:
    Set objEvent = objButton.Events(1)
    Set objVBScript = objButton.Events(1).Actions.AddAction(hmiActionCreationTypeVBScript)
    Select Case objVBScript.ScriptType
        Case 0
            strScriptType = "VB-Skript is used"
        Case 1
            strScriptType = "C-Skript is used"
    End Select
    MsgBox strScriptType
End Sub
