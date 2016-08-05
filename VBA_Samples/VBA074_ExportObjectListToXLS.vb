Sub ExportObjectListToXLS()
'VBA74
    Dim objGDApplicationApplication As grafexe.Application
    Dim objDoc As grafexe.Document
    Dim objHMIObject As grafexe.HMIObject
    Dim objProperty As grafexe.HMIProperty
    Dim objXLS As Excel.Application
    Dim objWSheet As Excel.Worksheet
    Dim objWBook As Excel.Workbook
    Dim lRow As Long
  
    'Define local errorhandler
    On Local Error GoTo LocErrTrap
  
    'Set references on the applications Excel and GraphicsDesigner
    Set objGDApplication = Application
    Set objDoc = objGDApplication.ActiveDocument
    Set objXLS = New Excel.Application
  
    'Create workbook
    Set objWBook = objXLS.Workbooks.Add()
    objWBook.SaveAs objGDApplication.ApplicationDataPath & "Export.xls"
  
    'Create worksheet in the new workbook and write headline
    'The name of the worksheet is equivalent to the documents name
    Set objWSheet = objWBook.Worksheets.Add
    objWSheet.Name = objDoc.Name
    objWSheet.Cells(1, 1) = "Objektname"
    objWSheet.Cells(1, 2) = "Objekttyp"
    objWSheet.Cells(1, 3) = "ProgID"
    objWSheet.Cells(1, 4) = "Position X"
    objWSheet.Cells(1, 5) = "Position Y"
    objWSheet.Cells(1, 6) = "Breite"
    objWSheet.Cells(1, 7) = "Höhe"
    objWSheet.Cells(1, 8) = "Ebene"
    lRow = 3
 
    'Every object will be written with their objectproperties width,
    'height, pos x, pos y and layer to Excel. If the object is an
    'ActiveX-Control the ProgID will be also exported.
    For Each objHMIObject In objDoc.HMIObjects
        DoEvents
        objWSheet.Cells(lRow, 1).value = objHMIObject.ObjectName
        objWSheet.Cells(lRow, 2).value = objHMIObject.Type
        If UCase(objHMIObject.Type) = "HMIACTIVEXCONTROL" Then
            objWSheet.Cells(lRow, 3).value = objHMIObject.ProgID
        End If
        objWSheet.Cells(lRow, 4).value = objHMIObject.Left
        objWSheet.Cells(lRow, 5).value = objHMIObject.Top
        objWSheet.Cells(lRow, 6).value = objHMIObject.Width
        objWSheet.Cells(lRow, 7).value = objHMIObject.Height
        objWSheet.Cells(lRow, 8).value = objHMIObject.Layer
        lRow = lRow + 1
    Next objHMIObject
    objWSheet.Columns.AutoFit
  
    Set objWSheet = Nothing
    objWBook.Save
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