VBS109
Dim cboComboBox
Set cboComboBox = ScreenItems("ComboBox1")
cboCombobox.AddItem "1_ComboBox_Field"
cboComboBox.AddItem "2_ComboBox_Field"
cboComboBox.AddItem "3_ComboBox_Field"
cboComboBox.FontBold = True
cboComboBox.FontItalic = True
cboComboBox.ListIndex = 2
