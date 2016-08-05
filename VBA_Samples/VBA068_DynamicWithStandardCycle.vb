Sub DynamicWithStandardCycle()
'VBA68
    Dim objVBScript As HMIScriptInfo
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_Standard", "HMICircle")
    Set objVBScript = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeVBScript)
    With objVBScript
        .Trigger.Type = hmiTriggerTypeStandardCycle
        '"CycleType"-specification is necessary:
        .Trigger.CycleType = hmiCycleType_10s
        .Trigger.Name = "VBA_StandardCycle"
        .SourceCode = ""
    End With
End Sub