Sub AddActionToPropertyTypeCScript()
'VBA251
    Dim objEvent As HMIEvent
    Dim objCScript As HMIScriptInfo
    Dim objCircle As HMICircle
    'Create circle in the picture. If property "Radius" is changed,
    'a C-action is added:
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_AB", "HMICircle")
    Set objEvent = objCircle.Radius.Events(1)
    Set objCScript = objEvent.Actions.AddAction(hmiActionCreationTypeCScript)
End Sub
