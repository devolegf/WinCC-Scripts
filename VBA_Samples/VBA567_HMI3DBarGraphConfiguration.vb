Sub HMI3DBarGraphConfiguration()
'VBA567
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer04Checked = True
        .Layer04Color = RGB(255, 0, 255)
    End With
End Sub
