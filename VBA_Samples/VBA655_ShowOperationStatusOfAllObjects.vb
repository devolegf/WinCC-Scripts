Sub ShowOperationStatusOfAllObjects()
'VBA655
    Dim objObject As HMIObject
    Dim bStatus As Boolean
    Dim strStatus As String
    Dim strName As String
    Dim iMax As Integer
    Dim iIndex As Integer
    Dim iAnswer As Integer
    iMax = ActiveDocument.HMIObjects.Count
    iIndex = 1
    For iIndex = 1 To iMax
        strName = ActiveDocument.HMIObjects(iIndex).ObjectName
        bStatus = ActiveDocument.HMIObjects(iIndex).Operation
        Select Case bStatus
            Case True
                strStatus = "yes"
            Case False
                strStatus = "no"
        End Select
        iAnswer = MsgBox("Object: " & strName & vbCrLf & "Operator-Control enable: " & strStatus, vbOKCancel)
        If vbCancel = iAnswer Then Exit For
    Next iIndex
    If 0 = iMax Then MsgBox "No objects in the active document."
End Sub
