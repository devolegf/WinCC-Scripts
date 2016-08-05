Sub DynamicWithWindowCycle()
'VBA71
    Dim objVBScript As HMIScriptInfo
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_Window", "HMICircle")
    Set objVBScript = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeVBScript)
    With objVBScript
        .Trigger.Type = hmiTriggerTypeWindowCycle
        .Trigger.Name = "VBA_WindowCycle"
        .SourceCode = ""
    End With
End Sub