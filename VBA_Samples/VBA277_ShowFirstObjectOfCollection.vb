Sub ShowFirstObjectOfCollection()
'VBA277
    Dim strName As String
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button", "HMIButton")
    strName = objButton.LDFonts(1).Family
    MsgBox strName
End Sub
