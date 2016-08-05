Sub AddStatusTextsToAppMenu1()
'VBA8
    Dim objMenu1 As HMIMenu
'
    'Object "objStatusTextMenuItem1" contains foreign-language texts
    Dim objStatusTextMenuItem1 As HMILanguageText
 
    Set objMenu1 = Application.CustomMenus("AppMenu1")
'
    'Assign a statustext to a menuitem:
    objMenu1.MenuItems("mItem1_1").StatusText = "Statustext of the first menuitem"
'
    'Assign a foreign statustext to a menuitem:
    Set objStatusTextMenuItem1 = objMenu1.MenuItems("mItem1_1").LDStatusTexts.Add(1033, "This is my first statustext in english")
End Sub