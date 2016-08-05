Sub AddDynamicAsVBSkriptToProperty()
'VBA357
    Dim objVBScript As HMIScriptInfo
    Dim objCircle As HMICircle
     
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle1", "HMICircle")
    Set objVBScript = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeVBScript)
'
    'Define cycletime and sourcecode
    With objVBScript
        .SourceCode = ""
        .Trigger.Type = hmiTriggerTypeStandardCycle
        .Trigger.CycleType = hmiCycleType_2s
        .Trigger.Name = "Trigger1"
    End With
End Sub
