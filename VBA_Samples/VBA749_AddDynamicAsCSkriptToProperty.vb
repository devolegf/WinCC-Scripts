Sub AddDynamicAsCSkriptToProperty()
'VBA749
    Dim objVBScript As HMIScriptInfo
    Dim objCircle As HMICircle
 
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("myCircle", "HMICircle")
    Set objVBScript = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeVBScript)
    With objVBScript
        .Trigger.Type = hmiTriggerTypeStandardCycle
        .Trigger.CycleType = hmiCycleType_2s
        .Trigger.Name = "Trigger1"
    End With
End Sub
