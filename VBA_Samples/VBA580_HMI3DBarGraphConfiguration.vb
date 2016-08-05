Sub HMI3DBarGraphConfiguration()
'VBA580
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer08Checked = True
        .Layer08Value = 80
    End With
End Sub
