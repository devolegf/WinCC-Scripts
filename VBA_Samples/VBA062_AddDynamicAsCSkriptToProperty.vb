Sub AddDynamicAsCSkriptToProperty()
'VBA62
    Dim objVBScript As HMIScriptInfo
    Dim objCircle As HMICircle
    Dim strCode As String
    strCode = "Dim myCircle" & vbCrLf & "Set myCircle = "
    strCode = strCode & "HMIRuntime.ActiveScreen.ScreenItems(""myCircle"")"
    strCode = strCode & vbCrLf & "Set myCircle.Radius = myCircle.Radius + 5"
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_VB", "HMICircle")
'
    'Create dynamic of property "Radius":
    Set objVBScript = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeVBScript)
'
    'Set SourceCode and cycletime:
    With objVBScript
        .SourceCode = strCode
        .Trigger.Type = hmiTriggerTypeStandardCycle
        .Trigger.CycleType = hmiCycleType_2s
        .Trigger.Name = "Trigger1"
    End With
End Sub
