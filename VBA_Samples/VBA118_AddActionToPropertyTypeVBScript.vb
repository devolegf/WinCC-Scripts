Sub AddActionToPropertyTypeVBScript()
'VBA118
    Dim objEvent As HMIEvent
    Dim objVBScript As HMIScriptInfo
    Dim objCircle As HMICircle
    'Create circle in picture. By changing of property "Radius"
    'a VBS-action will be started:
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_AB", "HMICircle")
    Set objEvent = objCircle.Radius.Events(1)
    Set objVBScript = objEvent.Actions.AddAction(hmiActionCreationTypeVBScript)
End Sub
