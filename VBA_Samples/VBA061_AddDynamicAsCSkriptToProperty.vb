Sub AddDynamicAsCSkriptToProperty()
'VBA61
    Dim objCScript As HMIScriptInfo
    Dim objCircle As HMICircle
    Dim strCode As String
    strCode = "long lHeight;" & vbCrLf & "int check;" & vbCrLf
    strCode = strCode & "GetHeight (""events.PDL"", ""myCircle""); & vbcrlf"
    strCode = strCode & "lHeight = lHeight+5;" & vbCrLf
    strCode = strCode & "check = SetHeight(""events.PDL"", ""myCircle"",lHeight);"
    strCode = strCode & vbCrLf & "//Return-Type: BOOL" & vbCrLf
    strCode = strCode & "return check;"
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_C", "HMICircle")
    'Create dynamic for Property "Radius":
    Set objCScript = objCircle.Height.CreateDynamic(hmiDynamicCreationTypeCScript)
'
    'set Sourcecode and cycletime:
    With objCScript
        .SourceCode = strCode
        .Trigger.Type = hmiTriggerTypeStandardCycle
        .Trigger.CycleType = hmiCycleType_2s
        .Trigger.Name = "Trigger1"
    End With
End Sub