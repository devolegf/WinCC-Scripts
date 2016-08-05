Sub PolygonCoordinatesOutput()
'VBA368
    Dim objPolyline As HMIPolyLine
    Dim iPosX As Integer
    Dim iPosY As Integer
    Dim iCounter As Integer
    Dim strResult As String
    iCounter = 1
    Set objPolyline = ActiveDocument.HMIObjects.AddHMIObject("Polyline1", "HMIPolyLine")
    For iCounter = 1 To objPolyline.PointCount
        With objPolyline
            .index = iCounter
            iPosX = .ActualPointLeft
            iPosY = .ActualPointTop
        End With
        strResult = strResult & vbCrLf & "Corner " & iCounter & ": x=" & iPosX & " y=" & iPosY
    Next iCounter
    MsgBox strResult
End Sub
