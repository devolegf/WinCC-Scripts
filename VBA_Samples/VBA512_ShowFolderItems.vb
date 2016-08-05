Sub ShowFolderItems()
'VBA512
    Dim colFolderItems As HMIFolderItems
    Dim objFolderItem As HMIFolderItem
    Dim iAnswer As Integer
    Dim iMaxFolder As Integer
    Dim iMaxSymbolLib As Integer
    Dim iSymbolLibIndex As Integer
    Dim iSubFolderIndex As Integer
    Dim strSubFolderName As String
    Dim strFolderItemName As String

    'determine the number of symbollibraries:
    iMaxSymbolLib = Application.SymbolLibraries.Count
    iSymbolLibIndex = 1
    For iSymbolLibIndex = 1 To iMaxSymbolLib
        With Application.SymbolLibraries(iSymbolLibIndex)
            Set colFolderItems = .FolderItems
'
            'To determine the number of folders in actual symbollibrary:
            iMaxFolder = .FolderItems.Count
            MsgBox "Number of FolderItems in " & .Name & " : " & iMaxFolder
'
            'Output all subfoldernames from actual folder:
            For Each objFolderItem In colFolderItems
                iSubFolderIndex = 1
                For iSubFolderIndex = 1 To iMaxFolder
                    strFolderItemName = objFolderItem.DisplayName
                    If 0 <> objFolderItem.Folder.Count Then
                        strSubFolderName = objFolderItem.Folder(iSubFolderIndex).DisplayName
                        iAnswer = MsgBox("SymbolLibrary: " & .Name & vbCrLf & "act. Folder: " & strFolderItemName & vbCrLf & "act. Subfolder: " & strSubFolderName, vbOKCancel)
'
                        'If "Cancel" is clicked, continued with next FolderItem
                        If vbCancel = iAnswer Then
                            Exit For
                        End If
                    Else
                        MsgBox "There are no subfolders in " & objFolderItem.DisplayName
                        Exit For
                    End If
                Next iSubFolderIndex
            Next objFolderItem
        End With
    Next iSymbolLibIndex
End Sub
