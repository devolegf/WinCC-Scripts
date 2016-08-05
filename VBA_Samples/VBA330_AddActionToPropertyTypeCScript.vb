Sub AddActionToPropertyTypeCScript()
'VBA330
    Dim objEvent As HMIEvent
    Dim objCScript As HMIScriptInfo
    Dim objCircle As HMICircle
    'Add circle to picture. By changing of property "Radius"
    'a C-Aktion is initiated:
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_AB", "HMICircle")
    Set objEvent = objCircle.Radius.Events(1)
    Set objCScript = objEvent.Actions.AddAction(hmiActionCreationTypeCScript)
End Sub
