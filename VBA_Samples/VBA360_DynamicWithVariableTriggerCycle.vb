Sub DynamicWithVariableTriggerCycle()
'VBA360
    Dim objVBScript As HMIScriptInfo
    Dim objVarTrigger As HMIVariableTrigger
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_VariableTrigger", "HMICircle")
    Set objVBScript = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeVBScript)
    With objVBScript
        'Definition of triggername and cycletime is to do with the Add-methode
        Set objVarTrigger = .Trigger.VariableTriggers.Add("VarTrigger", hmiVariableCycleType_10s)
        .SourceCode = ""
    End With
End Sub
