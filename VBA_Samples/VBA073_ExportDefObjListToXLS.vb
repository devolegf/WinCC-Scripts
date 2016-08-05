Sub ExportDefObjListToXLS()
'VBA73
    Dim objGDApplication As grafexe.Application
    Dim objHMIObject As grafexe.HMIObject
    Dim objProperty As grafexe.HMIProperty
    Dim objXLS As Excel.Application
    Dim objWSheet As Excel.Worksheet
    Dim objWBook As Excel.Workbook
    Dim rngSelection As Excel.Range
    Dim lRow As Long
    Dim lRowGroupStart As Long

    'define local errorhandler
    On Local Error GoTo LocErrTrap

    'Set references to the applications Excel and GraphicsDesigner
    Set objGDApplication = Application
    Set objXLS = New Excel.Application
  
    'Create workbook
    Set objWBook = objXLS.Workbooks.Add()
    objWBook.SaveAs objGDApplication.ApplicationDataPath & "DefaultObjekte.xls"
  
    'Adds new worksheet to the new workbook
    Set objWSheet = objWBook.Worksheets.Add
    objWSheet.Name = "DefaultObjekte"
    lRow = 1
  
    'Every object of the DefaultHMIObjects-collection will be written
    'to the worksheet with their objectproperties.
    'For better overview the objects will be grouped.
    For Each objHMIObject In objGDApplication.DefaultHMIObjects
        DoEvents
        objWSheet.Cells(lRow, 1).value = objHMIObject.ObjectName
        objWSheet.Cells(lRow, 2).value = objHMIObject.Type
        lRow = lRow + 1
    
        lRowGroupStart = lRow
        For Each objProperty In objHMIObject.Properties
            'Write displayed name and automationname of property
            'into the worksheet
            objWSheet.Cells(lRow, 2).value = objProperty.DisplayName
            objWSheet.Cells(lRow, 3).value = objProperty.Name
      
            'Write the value of property, datatype and if their dynamicable
            'into the worksheet
            If Not IsEmpty(objProperty.value) Then _
                        objWSheet.Cells(lRow, 4).value = objProperty.value
                objWSheet.Cells(lRow, 5).value = objProperty.IsDynamicable
                objWSheet.Cells(lRow, 6).value = TypeName(objProperty.value)
                objWSheet.Cells(lRow, 7).value = VarType(objProperty.value)
                lRow = lRow + 1
        Next objProperty
    
        'Select and groups the range of object-properties in the worksheet
        Set rngSelection = objWSheet.Range(objWSheet.Rows(lRowGroupStart), _
                                    objWSheet.Rows(lRow - 1))
        rngSelection.Select
        rngSelection.Group
        Set rngSelection = Nothing
    
        'Insert empty row
        lRow = lRow + 1
    Next objHMIObject
    
    objWSheet.Columns.AutoFit
  
    Set objWSheet = Nothing
    objWBook.Save
    objWBook.Close
    Set objWBook = Nothing
    objXLS.Quit
    Set objXLS = Nothing
    Set objGDApplication = Nothing
Exit Sub

LocErrTrap:
    MsgBox Err.Description, , Err.Source
    Resume Next
End Sub