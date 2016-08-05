Sub ImportObjectListFromXLS()
'VBA75
    Dim objGDApplication As grafexe.Application
    Dim objDoc As grafexe.Document
    Dim objHMIObject As grafexe.HMIObject
    Dim objXLS As Excel.Application
    Dim objWSheet As Excel.Worksheet
    Dim objWBook As Excel.Workbook
    Dim lRow As Long
    Dim strWorkbookName As String
    Dim strWorksheetName As String
    Dim strSheets As String
  
    'define local errorhandler
    On Local Error GoTo LocErrTrap
  
    'Set references on the applications Excel and GraphicsDesigner
    Set objGDApplication = Application
    Set objDoc = objGDApplication.ActiveDocument
    Set objXLS = New Excel.Application
  
  
    'Open workbook. The workbook have to be in datapath of GraphicsDesigner
    strWorkbookName = InputBox("Name of workbook:", "Import of objects")
    Set objWBook = objXLS.Workbooks.Open(objGDApplication.ApplicationDataPath & strWorkbookName)
    If objWBook Is Nothing Then
        MsgBox "Open workbook fails!" & vbCrLf & "This function is cancled!", vbCritical, "Import od objects"
        Set objDoc = Nothing
        Set objGDApplication = Nothing
        Set objXLS = Nothing
        Exit Sub
    End If
  
    'Read out the names of all worksheets contained in the workbook
    For Each objWSheet In objWBook.Sheets
        strSheets = strSheets & objWSheet.Name & vbCrLf
    Next objWSheet
    strWorksheetName = InputBox("Name of table to import:" & vbCrLf & strSheets, "Import of objects")
    Set objWSheet = objWBook.Sheets(strWorksheetName)
    lRow = 3
  
    'Import the worksheet as long as in actual row the first column is empty.
    'Add with the outreaded data new objects to the active document and
    'assign the values to the objectproperties
    With objWSheet
        While (.Cells(lRow, 1).value <> vbNullString) And (Not IsEmpty(.Cells(lRow, 1).value))
    
            'Add the objects to the document as its objecttype,
            'do nothing by groups, their have to create before.
            If (UCase(.Cells(lRow, 2).value) = "HMIGROUP") Then
    
            Else
                If (UCase(.Cells(lRow, 2).value) = "HMIACTIVEXCONTROL") Then
                    Set objHMIObject = objDoc.HMIObjects.AddActiveXControl(.Cells(lRow, 1).value, .Cells(lRow, 3).value)
                Else
                    Set objHMIObject = objDoc.HMIObjects.AddHMIObject(.Cells(lRow, 1).value, .Cells(lRow, 2).value)
                End If
                objHMIObject.Left = .Cells(lRow, 4).value
                objHMIObject.Top = .Cells(lRow, 5).value
                objHMIObject.Width = .Cells(lRow, 6).value
                objHMIObject.Height = .Cells(lRow, 7).value
                objHMIObject.Layer = .Cells(lRow, 8).value
            End If
  
            Set objHMIObject = Nothing
            lRow = lRow + 1
        Wend
    End With
    objWBook.Close
    Set objWBook = Nothing
    objXLS.Quit
    Set objXLS = Nothing
    Set objDoc = Nothing
    Set objGDApplication = Nothing
Exit Sub

LocErrTrap:
    MsgBox Err.Description, , Err.Source
    Resume Next
End Sub
