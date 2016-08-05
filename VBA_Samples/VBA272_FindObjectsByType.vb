Sub FindObjectsByType()
'VBA272
    Dim colSearchResults As HMICollection
    Dim objMember As HMIObject
    Dim iResult As Integer
    Dim strName As String
    Set colSearchResults = ActiveDocument.HMIObjects.Find(ObjectType:="HMICircle")
    For Each objMember In colSearchResults
        iResult = colSearchResults.Count
        strName = objMember.ObjectName
        MsgBox "Found: " & CStr(iResult) & vbCrLf & "Objectname: " & strName
    Next objMember
End Sub
