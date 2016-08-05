Sub GetVarName()
'VBA788
    Dim objVBScript As HMIScriptInfo
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.Item("Circle_VariableTrigger")
    Set objVBScript = objCircle.Radius.Dynamic
    With objVBScript
        'Reading out of variablename
        MsgBox "The radius is dynamicabled with: " & .Trigger.VariableTriggers.Item(1).VarName
    End With
End Sub
