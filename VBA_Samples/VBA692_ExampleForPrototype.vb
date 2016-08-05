Sub ExampleForPrototype()
'VBA692
    Dim objButton As HMIButton
    Dim objCircleA As HMICircle
    Dim objEvent As HMIEvent
    Dim objVBScript As HMIScriptInfo
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
    MsgBox objVBScript.Prototype
End Sub
