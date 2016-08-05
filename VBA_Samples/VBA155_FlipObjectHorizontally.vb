Sub FlipObjectHorizontally()
'VBA155
    Dim objStaticText As HMIStaticText
    Dim strPropertyName As String
    Dim iPropertyValue As Integer
    Set objStaticText = ActiveDocument.HMIObjects.AddHMIObject("Textfield", "HMIStaticText")
    strPropertyName = objStaticText.Properties("Text").Name
    With objStaticText
        .Width = 120
        .Text = "Sample Text"
        .Selected = True
        iPropertyValue = .AlignmentTop
        MsgBox "Value of '" & strPropertyName & "' before flip: " & iPropertyValue
        ActiveDocument.Selection.FlipHorizontally
        iPropertyValue = objStaticText.AlignmentTop
        MsgBox "Value of '" & strPropertyName & "' after flip: " & iPropertyValue
    End With
End Sub
