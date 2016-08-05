Sub AddDynamicAsCSkriptToProperty()
'VBA329
    Dim objCScript As HMIScriptInfo
    Dim objCircle As HMICircle
 
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle1", "HMICircle")
    Set objCScript = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeCScript)
'
    'Define triggertype and cycletime:
    With objCScript
        .SourceCode = ""
        .Trigger.Type = hmiTriggerTypeStandardCycle
        .Trigger.CycleType = hmiCycleType_2s
        .Trigger.Name = "Trigger1"
    End With
End Sub
