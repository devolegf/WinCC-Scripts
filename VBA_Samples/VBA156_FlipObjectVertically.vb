Sub FlipObjectVertically()
'VBA156
    Dim objStaticText As HMIStaticText
    Dim strPropertyName As String
    Dim iPropertyValue As Integer
    Set objStaticText = ActiveDocument.HMIObjects.AddHMIObject("Textfield", "HMIStaticText")
    strPropertyName = objStaticText.Properties("Text").Name
    With objStaticText
        .Width = 120
        .Text = "Sample Text"
        .Selected = True
        .AlignmentLeft = 0
        iPropertyValue = .AlignmentLeft
        MsgBox "Value of '" & strPropertyName & "' before flip: " & iPropertyValue
        ActiveDocument.Selection.FlipVertically
        iPropertyValue = objStaticText.AlignmentLeft
        MsgBox "Value of '" & strPropertyName & "' after flip: " & iPropertyValue
    End With
End Sub
