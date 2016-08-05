Sub PolyLineCoordsOutput()
'VBA681
    Dim iPcIndex As Integer
    Dim iPosX As Integer
    Dim iPosY As Integer
    Dim iIndex As Integer
    Dim objPolyLine As HMIPolyLine
    Set objPolyLine = Application.ActiveDocument.HMIObjects.AddHMIObject("PolyLine1", "HMIPolyLine")
    
'
    'Determine number of corners from "PolyLine1":
    iPcIndex = objPolyLine.PointCount
'
    'Output of x/y-coordinates from every corner:
    For iIndex = 1 To iPcIndex
        With objPolyLine
            .index = iIndex
            iPosX = .ActualPointLeft
            iPosY = .ActualPointTop
            MsgBox iIndex & ". corner:" & vbCrLf & "x-coordinate: " & iPosX & vbCrLf & "y-coordinate: " & iPosY
        End With
    Next iIndex
End Sub
