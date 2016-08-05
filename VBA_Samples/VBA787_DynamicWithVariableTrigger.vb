Sub DynamicWithVariableTrigger()
'VBA787
    Dim objVBScript As HMIScriptInfo
    Dim objVarTrigger As HMIVariableTrigger
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_VariableTrigger", "HMICircle")
    Set objVBScript = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeVBScript)
    With objVBScript
        'Triggername and cycletime are defined by add-methode
        Set objVarTrigger = .Trigger.VariableTriggers.Add("VarTrigger", hmiVariableCycleType_10s)
        .SourceCode = ""
    End With
End Sub
