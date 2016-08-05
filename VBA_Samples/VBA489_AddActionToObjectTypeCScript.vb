Sub AddActionToObjectTypeCScript()
'VBA489
    Dim objEvent As HMIEvent
    Dim objCScript As HMIScriptInfo
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_AB", "HMICircle")
'
    'C-action is initiated by click on object circle
    Set objEvent = objCircle.Events(1)
    Set objCScript = objEvent.Actions.AddAction(hmiActionCreationTypeCScript)
    MsgBox "the type of the projected event is " & objEvent.EventType
End Sub
