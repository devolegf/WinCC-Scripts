Sub AddActionToPropertyTypeCScript()
'VBA64
    Dim objEvent As HMIEvent
    Dim objCScript As HMIScriptInfo
    Dim objCircle As HMICircle
    'Create circle. Changing of the Property
    '"Radius" should be activate C-Aktion:
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_AB", "HMICircle")
    Set objEvent = objCircle.Radius.Events(1)
    Set objCScript = objEvent.Actions.AddAction(hmiActionCreationTypeCScript)
'
    'Assign a corresponding custom-function to the property "SourceCode":
    objCScript.SourceCode = ""
End Sub