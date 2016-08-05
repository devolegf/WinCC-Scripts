Sub ShowAnalogResultInfosOfCircleRadius()
'VBA207
    Dim colAResultInfos As HMIAnalogResultInfos
    Dim objAResultInfo As HMIAnalogResultInfo
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Dim iAnswer As Integer
    Dim varRange As Variant
    Dim varValue As Variant
    Set objCircle = ActiveDocument.HMIObjects("Circle_A")
    Set objDynDialog = objCircle.Radius.Dynamic
    Set colAResultInfos = objDynDialog.AnalogResultInfos
    For Each objAResultInfo In colAResultInfos
        varRange = objAResultInfo.RangeTo
        varValue = objAResultInfo.value
        iAnswer = MsgBox("Ranges of values from Circle_A-Radius:" & vbCrLf & "Range of value to: " & varRange & vbCrLf & "Value of property: " & varValue, vbOKCancel)
        If vbCancel = iAnswer Then Exit For
    Next objAResultInfo
End Sub
