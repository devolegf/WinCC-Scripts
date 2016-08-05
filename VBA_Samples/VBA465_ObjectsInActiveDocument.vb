Sub ObjectsInActiveDocument()
'VBA465
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim objDocument As Document
    Set objDocument = Application.Documents.Add hmiDocumentTypeVisible
    Dim iIndex As Integer
    iIndex = 1
    For iIndex = 1 To 5
        Set objCircle = objDocument.HMIObjects.AddHMIObject("Circle" & iIndex, "HMICircle")
        Set objRectangle = objDocument.HMIObjects.AddHMIObject("Rectangle" & iIndex, "HMIRectangle")
        With objCircle
            .Top = (10 * iIndex)
            .Left = (10 * iIndex)
        End With
        With objRectangle
            .Top = ((10 * iIndex) + 50)
            .Left = (10 * iIndex)
        End With
    Next iIndex
    MsgBox "There are " & objDocument.HMIObjects.Count & " objects in the document"
End Sub
