Sub ShowEventsOfAllObjectsInActiveDocument()
'VBA252
    Dim colEvents As HMIEvents
    Dim objEvent As HMIEvent
    Dim iMax As Integer
    Dim iIndex As Integer
    Dim iAnswer As Integer
    Dim strEventName As String
    Dim strObjectName As String
    Dim varEventType As Variant
    iIndex = 1
    iMax = ActiveDocument.HMIObjects.Count
    For iIndex = 1 To iMax
        Set colEvents = ActiveDocument.HMIObjects(iIndex).Events
        strObjectName = ActiveDocument.HMIObjects(iIndex).ObjectName
        For Each objEvent In colEvents
            strEventName = objEvent.EventName
            varEventType = objEvent.EventType
            iAnswer = MsgBox("Objectname: " & strObjectName & vbCrLf & "Eventtype: " & varEventType & vbCrLf & "Eventname: " & strEventName, vbOKCancel)
            If vbCancel = iAnswer Then Exit For
        Next objEvent
        If vbCancel = iAnswer Then Exit For
    Next iIndex
End Sub
