Sub DynamicWithVariableTriggerCycle()
'VBA69
    Dim objVBScript As HMIScriptInfo
    Dim objVarTrigger As HMIVariableTrigger
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_VariableTrigger", "HMICircle")
    Set objVBScript = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeVBScript)
    With objVBScript
        Set objVarTrigger = .Trigger.VariableTriggers.Add("VarTrigger", hmiVariableCycleType_10s)
        .SourceCode = ""
    End With
End Sub