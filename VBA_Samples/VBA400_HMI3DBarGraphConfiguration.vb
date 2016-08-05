Sub HMI3DBarGraphConfiguration()
'VBA400
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Background = False
    End With
End Sub
