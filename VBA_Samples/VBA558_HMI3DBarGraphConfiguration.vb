Sub HMI3DBarGraphConfiguration()
'VBA558
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer01Checked = True
        .Layer01Color = RGB(255, 0, 255)
    End With
End Sub
