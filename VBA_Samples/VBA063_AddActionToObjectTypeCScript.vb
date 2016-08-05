Sub AddActionToObjectTypeCScript()
'VBA63
    Dim objEvent As HMIEvent
    Dim objCScript As HMIScriptInfo
    Dim objCircle As HMICircle
    'Create circle. Click on object executes an C-action
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_AB", "HMICircle")
    Set objEvent = objCircle.Events(1)
    Set objCScript = objEvent.Actions.AddAction(hmiActionCreationTypeCScript)
'
    'Assign a corresponding custom-function to the property "SourceCode":
    objCScript.SourceCode = ""
End Sub