Sub FindObjectsByName()
'VBA36
    Dim colSearchResults As HMICollection
    Dim objMember As HMIObject
    Dim iResult As Integer
    Dim strName As String
'
    'Wildcards (?, *) are allowed
    Set colSearchResults = ActiveDocument.HMIObjects.Find(ObjectName:="*Circle*")
    For Each objMember In colSearchResults
        iResult = colSearchResults.Count
        strName = objMember.ObjectName
        MsgBox "Found: " & CStr(iResult) & vbCrLf & "Objectname: " & strName
    Next objMember
End Sub