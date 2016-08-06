Sub ChangeCurrentDataLanguage()
'VBA1
    Application.CurrentDataLanguage = 1033
    MsgBox "The Data language has been changed to english"
    Application.CurrentDataLanguage = 1031
    MsgBox "The Data language has been changed to german"
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddLanguagesToButton()
'VBA2
    Dim objLabelText As HMILanguageText
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
'
    'Set defaultlabel:
    objButton.Text = "Default-Text"
'
    'Add english label:
    Set objLabelText = objButton.LDTexts.Add(1033, "English Text")
    'Add german label:
    Set objLabelText = objButton.LDTexts.Add(1031, "German Text")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateApplicationMenus()
'VBA3
    'Declaration of menus...:
    Dim objMenu1 As HMIMenu
    Dim objMenu2 As HMIMenu
    '
    'Add menus. Parameters are "Position", "Key" und "DefaultLabel":
    Set objMenu1 = Application.CustomMenus.InsertMenu(1, "AppMenu1", "App_Menu_1")
    Set objMenu2 = Application.CustomMenus.InsertMenu(2, "AppMenu2", "App_Menu_2")
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub InsertMenuItems()
'VBA4
    Dim objMenu1 As HMIMenu
    Dim objMenu2 As HMIMenu
    Dim objMenuItem1 As HMIMenuItem
    Dim objSubMenu1 As HMIMenuItem
    'Create Menu:
    Set objMenu1 = Application.CustomMenus.InsertMenu(1, "AppMenu1", "App_Menu_1")
    
    'Next lines add menu-items to userdefined menu.
    'Parameters are "Position", "Key" and DefaultLabel:
    Set objMenuItem1 = objMenu1.MenuItems.InsertMenuItem(1, "mItem1_1", "App_MenuItem_1")
    Set objMenuItem1 = objMenu1.MenuItems.InsertMenuItem(2, "mItem1_2", "App_MenuItem_2")
'
    'Adds seperator to menu ("Position", "Key")
    Set objMenuItem1 = objMenu1.MenuItems.InsertSeparator(3, "mItem1_3")
'
    'Adds a submenu into a userdefined menu
    Set objSubMenu1 = objMenu1.MenuItems.InsertSubMenu(4, "mItem1_4", "App_SubMenu_1")
 '
    'Adds a menu-item into a submenu
    Set objMenuItem1 = objSubMenu1.SubMenu.InsertMenuItem(5, "mItem1_5", "App_SubMenuItem_1")
    Set objMenuItem1 = objSubMenu1.SubMenu.InsertMenuItem(6, "mItem1_6", "App_SubMenuItem_2")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub InsertMenuItems()
'VBA5
    'Execute this procedure first
    Dim objMenu1 As HMIMenu
    Dim objMenu2 As HMIMenu
    Dim objMenuItem1 As HMIMenuItem
    Dim objSubMenu1 As HMIMenuItem
    'Add Menu:
    Set objMenu1 = Application.CustomMenus.InsertMenu(1, "AppMenu1", "App_Menu_1")
    
    'Next lines add menu-items to userdefined menu.
    'Parameters are "Position", "Key" and DefaultLabel:
    Set objMenuItem1 = objMenu1.MenuItems.InsertMenuItem(1, "mItem1_1", "App_MenuItem_1")
    Set objMenuItem1 = objMenu1.MenuItems.InsertMenuItem(2, "mItem1_2", "App_MenuItem_2")
'
    'Adds seperator to menu ("Position", "Key")
    Set objMenuItem1 = objMenu1.MenuItems.InsertSeparator(3, "mItem1_3")
'
    'Adds a submenu to a userdefined menu
    Set objSubMenu1 = objMenu1.MenuItems.InsertSubMenu(4, "mItem1_4", "App_SubMenu_1")
 '
    'Adds a menu-item to a submenu
    Set objMenuItem1 = objSubMenu1.SubMenu.InsertMenuItem(5, "mItem1_5", "App_SubMenuItem_1")
    Set objMenuItem1 = objSubMenu1.SubMenu.InsertMenuItem(6, "mItem1_6", "App_SubMenuItem_2")
End Sub

Sub MultipleLanguagesForAppMenu1()
    'Execute this procedure after execution of "InsertMenuItems()" 

    'Object "objLanguageTextMenu1" contains the
    'foreign-language labels for the menu
    Dim objLanguageTextMenu1 As HMILanguageText
'
    'Object "objLanguageTextMenu1Item" contains the
    'foreign-language labels for the menu-items
    Dim objLanguageTextMenuItem1 As HMILanguageText
 
    Dim objMenu As HMIMenu
    Dim objSubMenu1 As HMIMenuItem
 
    Set objMenu1 = Application.CustomMenus("AppMenu1")
    Set objSubMenu1 = Application.CustomMenus("AppMenu1").MenuItems("mItem1_4")
'
    'Ads foreign-language label into a menu:
    '("Add(LCID, DisplayName)"-Methode:
    Set objLanguageTextMenu1 = objMenu1.LDLabelTexts.Add(1033, "English_App_Menu_1")
'
    'Adds foreign-language label into a menuitem:
    Set objLanguageTextMenuItem1 = objMenu1.MenuItems("mItem1_1").LDLabelTexts.Add(1033, "My first menu item")
'
    'Adds a foreign-language label into a submenu:
    Set objLanguageTextMenuItem1 = objSubMenu1.SubMenu.Item("mItem1_5").LDLabelTexts.Add(1033, "My first submenu item")
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateApplicationToolbars()
'VBA6
    'Declare toolbar-objects...:
    Dim objToolbar1 As HMIToolbar
    Dim objToolbar2 As HMIToolbar
'
    'Add the toolbars with parameter "Key"
    Set objToolbar1 = Application.CustomToolbars.Add("AppToolbar1")
    Set objToolbar2 = Application.CustomToolbars.Add("AppToolbar2")
    
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub InsertToolbarItems()
'VBA7
    Dim objToolbar1 As HMIToolbar
    Dim objToolbarItem1 As HMIToolbarItem
'
    'Add a new toolbar:
    Set objToolbar1 = Application.CustomToolbars.Add("AppToolbar1")
    'Adds two toolbar-items to the toolbar
    '("InsertToolbarItem(Position, Key, DefaultToolTipText)"-Methode):
    Set objToolbarItem1 = objToolbar1.ToolbarItems.InsertToolbarItem(1, "tItem1_1", "First Symbol-Icon")
    Set objToolbarItem1 = objToolbar1.ToolbarItems.InsertToolbarItem(3, "tItem1_2", "Second Symbol-Icon")
'
    'Adds a seperator between the two toolbar-items
    '("InsertSeparator(Position, Key)"-Methode):
    Set objToolbarItem1 = objToolbar1.ToolbarItems.InsertSeparator(2, "tSeparator1_3")
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddStatusAndTooltipTextsToAppToolbar1()
'VBA9
    Dim objToolbar1 As HMIToolbar
'
    'Variable "StatusTextToolbarItem1" for foreign statustexts
    Dim objStatusTextToolbarItem1 As HMILanguageText
'
    'Variable "TooltipTextToolbarItem1 for foreign tooltiptexts
    Dim objTooltipTextToolbarItem1 As HMILanguageText
 
    Set objToolbar1 = Application.CustomToolbars("AppToolbar1")
'
    'Assign a statustext to a toolbaritem:
    objToolbar1.ToolbarItems("tItem1_1").StatusText = "Statustext für das erste Symbol-Icon"
'
    'Assign a foreign statustext to a toolbaritem:
    Set objStatusTextToolbarItem1 = objToolbar1.ToolbarItems("tItem1_1").LDStatusTexts.Add(1033, "This is my first status text in english")
'
    'Assign a foreign tooltiptext to a toolbaritem:
    Set objTooltipTextToolbarItem1 = objToolbar1.ToolbarItems("tItem1_1").LDTooltipTexts.Add(1033, "This is my first tooltip text in english")
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
'VBA10
'The next declaration has to be placed in the modul section
Dim WithEvents theApp As grafexe.Application


Private Sub SetApplication()
    'This procedure has to be executed (with "F5") first
    Set theApp = grafexe.Application
End Sub


Private Sub theApp_MenuItemClicked(ByVal MenuItem As IHMIMenuItem)
    Dim objClicked As HMIMenuItem
    Dim varMenuItemKey As Variant
    Set objClicked = MenuItem
'
    '"varMenuItemKey" contains the value of parameter "Key"
    'from clicked menu-item
    varMenuItemKey = objClicked.Key
    Select Case varMenuItemKey
        Case "mItem1_1"
            MsgBox "The first menuitem was clicked!"
    End Select
End Sub

Private Sub theApp_ToolbarItemClicked(ByVal ToolbarItem As IHMIToolbarItem)
    Dim objClicked As HMIToolbarItem
    Dim varToolbarItemKey As Variant
    Set objClicked = ToolbarItem
'
    '"varToolbarItemKey" contains the value of parameter "Key"
    'from clicked toolbar-item
    varToolbarItemKey = objClicked.Key
    Select Case varToolbarItemKey
        Case "tItem1_1"
            MsgBox "The first symbol-icon was clicked!"
    End Select
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateDocumentMenusUsingMacroProperty()
'VBA11
    Dim objDocMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
    Set objDocMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DocMenu1", "Doc_Menu_1")
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(1, "dmItem1_1", "First Menuitem")
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(2, "dmItem1_2", "Second Menuitem")
'
    'Assign a VBA-macro to every menu item
    With ActiveDocument.CustomMenus("DocMenu1")
        .MenuItems("dmItem1_1").Macro = "TestMacro1"
        .MenuItems("dmItem1_2").Macro = "TestMacro2"
    End With
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub TestMacro1()
'VBA12
    MsgBox "TestMacro1 is execute"
End Sub
 
Sub TestMacro2()
    MsgBox "TestMacro2 is execute"
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
VBA13
Application.ActiveDocument.CustomMenus(1).MenuItems(1).Enabled = False


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
VBA14
Application.ActiveDocument.CustomMenus(1).MenuItems(2).Checked = True


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
VBA15
Application.ActiveDocument.CustomMenus(1).MenuItems(3).Shortcut = "Strg+G"


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
VBA16
Application.ActiveDocument.CustomMenus(1).MenuItems(4).Visible = False


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
VBA17
Application.SymbolLibraries(1)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
VBA18
Application.SymbolLibraries(1).FolderItems("Folder2")


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
VBA19
Application.SymbolLibraries(1).FolderItems("Folder2").Folder("Folder2").Folder.Item("Object1").DisplayName


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddNewFolderToProjectLibrary()
'VBA20
    Dim objProjectLib As HMISymbolLibrary
    Set objProjectLib = Application.SymbolLibraries(2)
'
    '("AddFolder(DefaultName)"-Methode):
    objProjectLib.FolderItems.AddFolder ("Custom Folder")
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CopyObjectFromGlobalLibraryToProjectLibrary()
'VBA21
    Dim objGlobalLib As HMISymbolLibrary
    Dim objProjectLib As HMISymbolLibrary
    Set objGlobalLib = Application.SymbolLibraries(1)
    Set objProjectLib = Application.SymbolLibraries(2)
'
    'Copies object "PC" from the "Global Library" into the clipboard
    objGlobalLib.FolderItems("Folder2").Folder("Folder2").Folder.Item("Object1").CopyToClipboard
'
    'The folder "Custom Folder" has to be available
    objProjectLib.FolderItems("Folder1").Folder.AddFromClipBoard ("Copy of PC/PLC")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddObjectFromPictureToProjectLibrary()
'VBA22
    Dim objProjectLib As HMISymbolLibrary
    Dim objCircle As HMICircle
    Set objProjectLib = Application.SymbolLibraries(2)
'
    'Insert new object "Circle1"
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle1", "HMICircle")
'
    'The folder "Custom Folder" has to be available
    '("AddItem(DefaultName, pHMIObject)"-Methode):
    objProjectLib.FolderItems("Folder1").Folder.AddItem "ProjectLib Circle", ActiveDocument.HMIObjects("Circle1")
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DeleteObjectFromProjectLibrary()
'VBA23
    Dim objProjectLib As HMISymbolLibrary
    Set objProjectLib = Application.SymbolLibraries(2)
'
    'The folder "Custom Folder" has to be available
    '("Delete"-Methode):
    objProjectLib.FolderItems("Folder1").Delete
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CopyObjectFromGlobalLibraryToActiveDocument()
'VBA24
    Dim objGlobalLib As HMISymbolLibrary
    Dim objHMIObject As HMIObject
    Dim iLastObject As Integer
    Set objGlobalLib = Application.SymbolLibraries(1)
'
    'Copy object "PC" from "Global Library" to clipboard
    objGlobalLib.FolderItems("Folder2").Folder("Folder2").Folder.Item("Object1").CopyToClipboard
'
    'Get object from clipboard and add it to active document
    ActiveDocument.PasteClipboard
'
    'Get last inserted object
    iLastObject = ActiveDocument.HMIObjects.Count
    Set objHMIObject = ActiveDocument.HMIObjects(iLastObject)
'
    'Set position of the object:
    With objHMIObject
        .Left = 40
        .Top = 40
    End With
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateDocumentMenus()
'VBA25
    'Declare menuobjects:
    Dim objMenu1 As HMIMenu
    Dim objMenu2 As HMIMenu
    
    'Insert Menus ("InsertMenu"-Methode) with
    'Parameters - "Position", "Key", "DefaultLabel":
    Set objMenu1 = ActiveDocument.CustomMenus.InsertMenu(1, "DocMenu1", "Doc_Menu_1")
    Set objMenu2 = ActiveDocument.CustomMenus.InsertMenu(2, "DocMenu2", "Doc_Menu_2")
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateDocumentToolbars()
'VBA26
    'Declare toolbarobjects:
    Dim objToolbar1 As HMIToolbar
    Dim objToolbar2 As HMIToolbar
'
    'Insert toolbars ("Add"-Methode) with
    'Parameter - "Key":
    Set objToolbar1 = ActiveDocument.CustomToolbars.Add("DocToolbar1")
    Set objToolbar2 = ActiveDocument.CustomToolbars.Add("DocToolbar2")
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ConfigureSettingsOfLayer 
'VBA27
    Dim objLayer As HMILayer 
    Set objLayer = ActiveDocument.Layers(1) 
    With objLayer 
        'Configure "Layer 0"
        .MinZoom = 10 
        .MaxZoom = 100 
        .Name = "Configured with VBA" 
    End With
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateAndActivateView()
'VBA28
    Dim objView As HMIView
    Set objView = ActiveDocument.Views.Add
    objView.Activate
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SetZoomAndScrollPositionInActiveView()
'VBA29
    Dim objView As HMIView 
    Set objView = ActiveDocument.Views.Add 
    With objView 
        .Activate 
        .ScrollPosX = 40 
        .ScrollPosY = 10 
        .Zoom = 150
    End With
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddObject()
'VBA30
    Dim objObject As HMIObject
    Set objObject = ActiveDocument.HMIObjects.AddHMIObject("CircleAsHMIObject", "HMICircle")
'
    'standard-properties (e.g. the position) are available every time:
    objObject.Top = 40
    objObject.Left = 40
'
    'non-standard properties can be accessed using the Properties-collection:
    objObject.Properties("FlashBackColor") = True
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddCircle()
'VBA31
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("CircleAsHMICircle", "HMICircle")
'
    'The same as in example 1, but here you can set/get direct the
    specific properties of the circle:
    objCircle.Top = 80
    objCircle.Left = 80
    objCircle.FlashBackColor = True
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddCircle()
'VBA32
    'Creates object of type "HMICircle"
    Dim objCircle As HMICircle
'
    'Add object in active document
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("My Circle", "HMICircle")
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditDefinedObjectType()
'VBA33
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("myCircleAsCircle", "HMICircle")
    With objCircle
        'direct access of objectproperties available
        .BorderWidth = 4
        .BorderColor = RGB(255, 0, 255)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditHMIObject()
'VBA34
    Dim objObject As HMIObject
    Set objObject = ActiveDocument.HMIObjects.AddHMIObject("myCircleAsObject", "HMICircle")
    With objObject
        'Using object's properties collection to access its specific properties
        .Properties("BorderWidth") = 4
        .Properties("BorderColor") = RGB(255, 0, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SelectObject()
'VBA35
    Dim objObject As HMIObject
    Set objObject = ActiveDocument.HMIObjects.AddHMIObject("mySelectedCircle", "HMICircle")
    ActiveDocument.HMIObjects("mySelectedCircle").Selected = True
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub FindObjectsByName()
'VBA36
    Dim colSearchResults As HMICollection
    Dim objMember As HMIObject
    Dim iResult As Integer
    Dim strName As String
'
    'Wildcards (?, *) are allowed
    Set colSearchResults = ActiveDocument.HMIObjects.Find(ObjectName:="*Circle*")
    For Each objMember In colSearchResults
        iResult = colSearchResults.Count
        strName = objMember.ObjectName
        MsgBox "Found: " & CStr(iResult) & vbCrLf & "Objectname: " & strName
    Next objMember
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub FindObjectsByType()
'VBA37
    Dim colSearchResults As HMICollection
    Dim objMember As HMIObject
    Dim iResult As Integer
    Dim strName As String
    Set colSearchResults = ActiveDocument.HMIObjects.Find(ObjectType:="HMICircle")
    For Each objMember In colSearchResults
        iResult = colSearchResults.Count
        strName = objMember.ObjectName
        MsgBox "Found: " & CStr(iResult) & vbCrLf & "Objektname: " & strName
    Next objMember
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub FindObjectsByProperty()
'VBA38
    Dim colSearchResults As HMICollection
    Dim objMember As HMIObject
    Dim iResult As Integer
    Dim strName As String
    Set colSearchResults = ActiveDocument.HMIObjects.Find(PropertyName:="BackColor")
    For Each objMember In colSearchResults
        iResult = colSearchResults.Count
        strName = objMember.ObjectName
        MsgBox "Found: " & CStr(iResult) & vbCrLf & "Objectname: " & strName
    Next objMember
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DeleteObject()
'VBA39
    'Delete first object in active document:
    ActiveDocument.HMIObjects(1).Delete
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddOLEObjectToActiveDocument()
'VBA40
    Dim objOLEObject As HMIOLEObject
    Set objOLEObject = ActiveDocument.HMIObjects.AddOLEObject("MS Wordpad Document1", "Wordpad.Document.1")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddActiveXControl()
'VBA41
    Dim objActiveXControl As HMIActiveXControl
    Set objActiveXControl = ActiveDocument.HMIObjects.AddActiveXControl("WinCC_Gauge", "XGAUGE.XGaugeCtrl.1")
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddActiveXControl()
'VBA42
    Dim objActiveXControl As HMIActiveXControl
    Set objActiveXControl = ActiveDocument.HMIObjects.AddActiveXControl("WinCC_Gauge2", "XGAUGE.XGaugeCtrl.1")
'
    'move ActiveX-control:
    objActiveXControl.Top = 40
    objActiveXControl.Left = 60
'
    'Change individual property:
    objActiveXControl.Properties("BackColor").value = RGB(255, 0, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateGroup()
'VBA43
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim objGroup As HMIGroup
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    With objCircle
        .Top = 40
        .Left = 40
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 80
        .Selected = True
    End With
    MsgBox "Objects selected!"
    Set objGroup = ActiveDocument.Selection.CreateGroup
    objGroup.ObjectName = "myGroup"
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ModifyPropertyOfObjectInGroup()
'VBA44
    Dim objGroup As HMIGroup
    Set objGroup = ActiveDocument.HMIObjects("myGroup")
    objGroup.GroupedHMIObjects(1).Properties("BorderColor") = RGB(255, 0, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateGroup()
'VBA45
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim objGroup As HMIGroup
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    With objCircle
        .Top = 40
        .Left = 40
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 80
        .Selected = True
    End With
    MsgBox "Objects selected!"
    Set objGroup = ActiveDocument.Selection.CreateGroup
    'The name identifies the group-object
    objGroup.ObjectName = "My Group"
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddObjectToGroup()
'VBA46
    Dim objGroup As HMIGroup
    Dim objEllipseSegment As HMIEllipseSegment
    'Adds new object to active document
    Set objEllipseSegment = ActiveDocument.HMIObjects.AddHMIObject("EllipseSegment", "HMIEllipseSegment")
    Set objGroup = ActiveDocument.HMIObjects("My Group")
    'Adds the object to the group
    objGroup.GroupedHMIObjects.Add ("EllipseSegment")
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RemoveObjectFromGroup()
'VBA47
    Dim objGroup As HMIGroup
    Set objGroup = ActiveDocument.HMIObjects("My Group")
    'delete group-object' first object
    objGroup.GroupedHMIObjects.Remove (1)
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub UnGroup()
'VBA48
    Dim objGroup As HMIGroup
    Set objGroup = ActiveDocument.HMIObjects("My Group")
    objGroup.UnGroup
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DeleteGroup()
'VBA49
    Dim objGroup As HMIGroup
    Set objGroup = ActiveDocument.HMIObjects("My Group")
    objGroup.Delete
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ChangePropertiesOfGroupMembers()
'VBA50
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim objEllipse As HMIEllipse
    Dim objGroup As HMIGroup
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    Set objEllipse = ActiveDocument.HMIObjects.AddHMIObject("sEllipse", "HMIEllipse")
    With objCircle
        .Top = 40
        .Left = 40
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 80
        .Selected = True
        End With
    With objEllipse
        .Top = 120
        .Left = 120
        .Selected = True
    End With
    MsgBox "Objects selected!"
    Set objGroup = ActiveDocument.Selection.CreateGroup
    objGroup.ObjectName = "My Group"
    'Set bordercolor of 1. object = "red":
    objGroup.GroupedHMIObjects(1).Properties("BorderColor") = RGB(255, 0, 0)
    'set x-coordinate of 2. object = "120" :
    objGroup.GroupedHMIObjects(2).Properties("Left") = 120
    'set y-coordinate of 3. object = "90":
    objGroup.GroupedHMIObjects(3).Properties("Top") = 90
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ChangePropertiesOfAllGroupMembers()
'VBA51
    Dim objGroup As HMIGroup
    Dim iMaxMembers As Integer
    Dim iIndex As Integer
    Set objGroup = ActiveDocument.HMIObjects("My Group")
    iIndex = 1
'
    'Get number of objects in group-object:
    iMaxMembers = objGroup.GroupedHMIObjects.Count
'
    'set linecolor of all objects = "yellow":
    For iIndex = 1 To iMaxMembers
        objGroup.GroupedHMIObjects(iIndex).Properties("BorderColor") = RGB(255, 255, 0)
    Next iIndex
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateCustomizedObject()
'VBA52
    Dim objCustomizedObject As HMICustomizedObject
    Set objCustomizedObject = ActiveDocument.Selection.CreateCustomizedObject
    objCustomizedObject.ObjectName = "My Customized Object"
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditCustomizedObjectProperty()
'VBA53
    Dim objCustomizedObject As HMICustomizedObject
    Set objCustomizedObject = ActiveDocument.HMIObjects(1)
    objCustomizedObject.Properties("BackColor") = RGB(255, 0, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateCustomizedObject()
'VBA54
    Dim objCustomizedObject As HMICustomizedObject
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    With objCircle
        .Top = 40
        .Left = 40
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 80
        .Selected = True
    End With
    MsgBox "Objects selected!"
    Set objCustomizedObject = ActiveDocument.Selection.CreateCustomizedObject
'
    '*** The "Configurationdialog" started. ***
    '*** Configure the costumize-object with the "configurationdialog" ***
'
    objCustomizedObject.ObjectName = "My Customized Object"
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DestroyCustomizedObject()
'VBA55
    Dim objCustomizedObject As HMICustomizedObject
    Set objCustomizedObject = ActiveDocument.HMIObjects("My Customized Object")
    objCustomizedObject.Destroy
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DeleteCustomizedObject()
'VBA56
    Dim objCustomizedObject As HMICustomizedObject
    Set objCustomizedObject = ActiveDocument.HMIObjects("My Customized Object")
    objCustomizedObject.Delete
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateDynamicOnProperty()
'VBA57
    Dim objVariableTrigger As HMIVariableTrigger
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle1", "HMICircle")
'
    'Create dynamic with type "direct Variableconnection" at the
    'property "Radius":
    Set objVariableTrigger = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeVariableDirect, "NewDynamic1")
'
    'To complete dynamic, e.g. define cycle:
    With objVariableTrigger
        .CycleType = hmiVariableCycleType_2s
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicAsVariableDirectToProperty()
'VBA58
    Dim objVariableTrigger As HMIVariableTrigger
    Dim objCircle As HMICircle
 
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle1", "HMICircle")
    'Create dynamic at property "Top"
    Set objVariableTrigger = objCircle.Top.CreateDynamic(hmiDynamicCreationTypeVariableDirect, "NewDynamic1")
'
    'define cycle-time
    With objVariableTrigger
        .CycleType = hmiVariableCycleType_2s
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicAsVariableIndirectToProperty()
'VBA59
    Dim objVariableTrigger As HMIVariableTrigger
    Dim objCircle As HMICircle
 
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle2", "HMICircle")
    'Create dynamic on property "Left":
    Set objVariableTrigger = objCircle.Left.CreateDynamic(hmiDynamicCreationTypeVariableIndirect, "'NewDynamic1'")
'
    'Define cycle-time
    With objVariableTrigger
        .CycleType = hmiVariableCycleType_2s
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeAnalog()
'VBA60
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
'
    'Create dynamic
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
'
    'Configure dynamic. "ResultType" defines the type of valuerange:
    With objDynDialog
        .ResultType = hmiResultTypeAnalog
        .AnalogResultInfos.Add 50, 40
        .AnalogResultInfos.Add 100, 80
        .AnalogResultInfos.ElseCase = 100
    End With
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicAsCSkriptToProperty()
'VBA61
    Dim objCScript As HMIScriptInfo
    Dim objCircle As HMICircle
    Dim strCode As String
    strCode = "long lHeight;" & vbCrLf & "int check;" & vbCrLf
    strCode = strCode & "GetHeight (""events.PDL"", ""myCircle""); & vbcrlf"
    strCode = strCode & "lHeight = lHeight+5;" & vbCrLf
    strCode = strCode & "check = SetHeight(""events.PDL"", ""myCircle"",lHeight);"
    strCode = strCode & vbCrLf & "//Return-Type: BOOL" & vbCrLf
    strCode = strCode & "return check;"
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_C", "HMICircle")
    'Create dynamic for Property "Radius":
    Set objCScript = objCircle.Height.CreateDynamic(hmiDynamicCreationTypeCScript)
'
    'set Sourcecode and cycletime:
    With objCScript
        .SourceCode = strCode
        .Trigger.Type = hmiTriggerTypeStandardCycle
        .Trigger.CycleType = hmiCycleType_2s
        .Trigger.Name = "Trigger1"
    End With
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicAsCSkriptToProperty()
'VBA62
    Dim objVBScript As HMIScriptInfo
    Dim objCircle As HMICircle
    Dim strCode As String
    strCode = "Dim myCircle" & vbCrLf & "Set myCircle = "
    strCode = strCode & "HMIRuntime.ActiveScreen.ScreenItems(""myCircle"")"
    strCode = strCode & vbCrLf & "Set myCircle.Radius = myCircle.Radius + 5"
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_VB", "HMICircle")
'
    'Create dynamic of property "Radius":
    Set objVBScript = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeVBScript)
'
    'Set SourceCode and cycletime:
    With objVBScript
        .SourceCode = strCode
        .Trigger.Type = hmiTriggerTypeStandardCycle
        .Trigger.CycleType = hmiCycleType_2s
        .Trigger.Name = "Trigger1"
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddActionToObjectTypeCScript()
'VBA63
    Dim objEvent As HMIEvent
    Dim objCScript As HMIScriptInfo
    Dim objCircle As HMICircle
    'Create circle. Click on object executes an C-action
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_AB", "HMICircle")
    Set objEvent = objCircle.Events(1)
    Set objCScript = objEvent.Actions.AddAction(hmiActionCreationTypeCScript)
'
    'Assign a corresponding custom-function to the property "SourceCode":
    objCScript.SourceCode = ""
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddActionToPropertyTypeCScript()
'VBA64
    Dim objEvent As HMIEvent
    Dim objCScript As HMIScriptInfo
    Dim objCircle As HMICircle
    'Create circle. Changing of the Property
    '"Radius" should be activate C-Aktion:
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_AB", "HMICircle")
    Set objEvent = objCircle.Radius.Events(1)
    Set objCScript = objEvent.Actions.AddAction(hmiActionCreationTypeCScript)
'
    'Assign a corresponding custom-function to the property "SourceCode":
    objCScript.SourceCode = ""
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DirectConnection()
'VBA65
    Dim objButton As HMIButton
    Dim objRectangleA As HMIRectangle
    Dim objRectangleB As HMIRectangle
    Dim objEvent As HMIEvent
    Dim objDConnection As HMIDirectConnection
'
    'Create objects:
    Set objRectangleA = ActiveDocument.HMIObjects.AddHMIObject("Rectangle_A", "HMIRectangle")
    Set objRectangleB = ActiveDocument.HMIObjects.AddHMIObject("Rectangle_B", "HMIRectangle")
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    With objRectangleA
        .Top = 100
        .Left = 100
    End With
    With objRectangleB
        .Top = 250
        .Left = 400
        .BackColor = RGB(255, 0, 0)
    End With
    With objButton
        .Top = 10
        .Left = 10
        .Text = "SetPosition"
    End With
'
    'Directconnection is initiated by mouseclick:
    Set objDConnection = objButton.Events(1).Actions.AddAction(hmiActionCreationTypeDirectConnection)
    With objDConnection
        'Sourceobject: Property "Top" of Rectangle_A
        .SourceLink.Type = hmiSourceTypeProperty
        .SourceLink.ObjectName = "Rectangle_A"
        .SourceLink.AutomationName = "Top"
'
        'Destinationobject: Property "Left" of Rectangle_B
        .DestinationLink.Type = hmiDestTypeProperty
        .DestinationLink.ObjectName = "Rectangle_B"
        .DestinationLink.AutomationName = "Left"
    End With
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateVBActionToClickedEvent()
'VBA66
    Dim objButton As HMIButton
    Dim objCircle As HMICircle
    Dim objEvent As HMIEvent
    Dim objCScript As HMIScriptInfo
    Dim strCode As String
    strCode = "long lHeight;" & vbCrLf & "int check;" & vbCrLf
    strCode = strCode & "lHeight = GetHeight (""events.PDL"", ""myCircle"");"
    strCode = strCode & vbCrLf & "lHeight = lHeight+5;" & vbCrLf & "check = "
    strCode = strCode & "SetHeight(""events.PDL"", ""myCircle"",lHeight);"
    strCode = strCode & vbCrLf & "//Return-Type: Void"
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_C", "HMICircle")
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    With objCircle
        .Top = 100
        .Left = 100
        .BackColor = RGB(255, 0, 0)
    End With
    With objButton
        .Top = 10
        .Left = 10
        .Text = "Increase Radius"
    End With
    'Configure directconnection:
    Set objCScript = objButton.Events(1).Actions.AddAction(hmiActionCreationTypeCScript)
    With objCScript
'
        'Note: Replace "events.PDL" with your picturename
        .SourceCode = strCode
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateVBActionToClickedEvent()
'VBA67
    Dim objButton As HMIButton
    Dim objCircle As HMICircle
    Dim objEvent As HMIEvent
    Dim objVBScript As HMIScriptInfo
    Dim strCode As String
    strCode = "Dim myCircle" & vbCrLf & "Set myCircle = "
    strCode = strCode & "HMIRuntime.ActiveScreen.ScreenItems(""Circle_VB"")"
    strCode = strCode & vbCrLf & "myCircle.Radius = myCircle.Radius + 5"
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_VB", "HMICircle")
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    With objCircle
        .Top = 100
        .Left = 100
        .BackColor = RGB(255, 0, 0)
    End With
    With objButton
        .Top = 10
        .Left = 10
        .Width = 120
        .Text = "Increase Radius"
    End With
    'Define event and assign sourcecode:
    Set objVBScript = objButton.Events(1).Actions.AddAction(hmiActionCreationTypeVBScript)
    With objVBScript
        .SourceCode = strCode
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DynamicWithStandardCycle()
'VBA68
    Dim objVBScript As HMIScriptInfo
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_Standard", "HMICircle")
    Set objVBScript = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeVBScript)
    With objVBScript
        .Trigger.Type = hmiTriggerTypeStandardCycle
        '"CycleType"-specification is necessary:
        .Trigger.CycleType = hmiCycleType_10s
        .Trigger.Name = "VBA_StandardCycle"
        .SourceCode = ""
    End With
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DynamicWithVariableTriggerCycle()
'VBA69
    Dim objVBScript As HMIScriptInfo
    Dim objVarTrigger As HMIVariableTrigger
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_VariableTrigger", "HMICircle")
    Set objVBScript = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeVBScript)
    With objVBScript
        Set objVarTrigger = .Trigger.VariableTriggers.Add("VarTrigger", hmiVariableCycleType_10s)
        .SourceCode = ""
    End With
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DynamicWithPictureCycle()
'VBA70
    Dim objVBScript As HMIScriptInfo
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_Picture", "HMICircle")
    Set objVBScript = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeVBScript)
    With objVBScript
        .Trigger.Type = hmiTriggerTypePictureCycle
        .Trigger.Name = "VBA_PictureCycle"
        .SourceCode = ""
    End With
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DynamicWithWindowCycle()
'VBA71
    Dim objVBScript As HMIScriptInfo
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_Window", "HMICircle")
    Set objVBScript = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeVBScript)
    With objVBScript
        .Trigger.Type = hmiTriggerTypeWindowCycle
        .Trigger.Name = "VBA_WindowCycle"
        .SourceCode = ""
    End With
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub Document_HMIObjectPropertyChanged(ByVal Property As IHMIProperty, CancelForwarding As Boolean)
'VBA72
    CancelForwarding = True
    MsgBox "Object's property has been changed!"
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ExportDefObjListToXLS()
'VBA73
    Dim objGDApplication As grafexe.Application
    Dim objHMIObject As grafexe.HMIObject
    Dim objProperty As grafexe.HMIProperty
    Dim objXLS As Excel.Application
    Dim objWSheet As Excel.Worksheet
    Dim objWBook As Excel.Workbook
    Dim rngSelection As Excel.Range
    Dim lRow As Long
    Dim lRowGroupStart As Long

    'define local errorhandler
    On Local Error GoTo LocErrTrap

    'Set references to the applications Excel and GraphicsDesigner
    Set objGDApplication = Application
    Set objXLS = New Excel.Application
  
    'Create workbook
    Set objWBook = objXLS.Workbooks.Add()
    objWBook.SaveAs objGDApplication.ApplicationDataPath & "DefaultObjekte.xls"
  
    'Adds new worksheet to the new workbook
    Set objWSheet = objWBook.Worksheets.Add
    objWSheet.Name = "DefaultObjekte"
    lRow = 1
  
    'Every object of the DefaultHMIObjects-collection will be written
    'to the worksheet with their objectproperties.
    'For better overview the objects will be grouped.
    For Each objHMIObject In objGDApplication.DefaultHMIObjects
        DoEvents
        objWSheet.Cells(lRow, 1).value = objHMIObject.ObjectName
        objWSheet.Cells(lRow, 2).value = objHMIObject.Type
        lRow = lRow + 1
    
        lRowGroupStart = lRow
        For Each objProperty In objHMIObject.Properties
            'Write displayed name and automationname of property
            'into the worksheet
            objWSheet.Cells(lRow, 2).value = objProperty.DisplayName
            objWSheet.Cells(lRow, 3).value = objProperty.Name
      
            'Write the value of property, datatype and if their dynamicable
            'into the worksheet
            If Not IsEmpty(objProperty.value) Then _
                        objWSheet.Cells(lRow, 4).value = objProperty.value
                objWSheet.Cells(lRow, 5).value = objProperty.IsDynamicable
                objWSheet.Cells(lRow, 6).value = TypeName(objProperty.value)
                objWSheet.Cells(lRow, 7).value = VarType(objProperty.value)
                lRow = lRow + 1
        Next objProperty
    
        'Select and groups the range of object-properties in the worksheet
        Set rngSelection = objWSheet.Range(objWSheet.Rows(lRowGroupStart), _
                                    objWSheet.Rows(lRow - 1))
        rngSelection.Select
        rngSelection.Group
        Set rngSelection = Nothing
    
        'Insert empty row
        lRow = lRow + 1
    Next objHMIObject
    
    objWSheet.Columns.AutoFit
  
    Set objWSheet = Nothing
    objWBook.Save
    objWBook.Close
    Set objWBook = Nothing
    objXLS.Quit
    Set objXLS = Nothing
    Set objGDApplication = Nothing
Exit Sub

LocErrTrap:
    MsgBox Err.Description, , Err.Source
    Resume Next
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ImportObjectListFromXLS()
'VBA75
    Dim objGDApplication As grafexe.Application
    Dim objDoc As grafexe.Document
    Dim objHMIObject As grafexe.HMIObject
    Dim objXLS As Excel.Application
    Dim objWSheet As Excel.Worksheet
    Dim objWBook As Excel.Workbook
    Dim lRow As Long
    Dim strWorkbookName As String
    Dim strWorksheetName As String
    Dim strSheets As String
  
    'define local errorhandler
    On Local Error GoTo LocErrTrap
  
    'Set references on the applications Excel and GraphicsDesigner
    Set objGDApplication = Application
    Set objDoc = objGDApplication.ActiveDocument
    Set objXLS = New Excel.Application
  
  
    'Open workbook. The workbook have to be in datapath of GraphicsDesigner
    strWorkbookName = InputBox("Name of workbook:", "Import of objects")
    Set objWBook = objXLS.Workbooks.Open(objGDApplication.ApplicationDataPath & strWorkbookName)
    If objWBook Is Nothing Then
        MsgBox "Open workbook fails!" & vbCrLf & "This function is cancled!", vbCritical, "Import od objects"
        Set objDoc = Nothing
        Set objGDApplication = Nothing
        Set objXLS = Nothing
        Exit Sub
    End If
  
    'Read out the names of all worksheets contained in the workbook
    For Each objWSheet In objWBook.Sheets
        strSheets = strSheets & objWSheet.Name & vbCrLf
    Next objWSheet
    strWorksheetName = InputBox("Name of table to import:" & vbCrLf & strSheets, "Import of objects")
    Set objWSheet = objWBook.Sheets(strWorksheetName)
    lRow = 3
  
    'Import the worksheet as long as in actual row the first column is empty.
    'Add with the outreaded data new objects to the active document and
    'assign the values to the objectproperties
    With objWSheet
        While (.Cells(lRow, 1).value <> vbNullString) And (Not IsEmpty(.Cells(lRow, 1).value))
    
            'Add the objects to the document as its objecttype,
            'do nothing by groups, their have to create before.
            If (UCase(.Cells(lRow, 2).value) = "HMIGROUP") Then
    
            Else
                If (UCase(.Cells(lRow, 2).value) = "HMIACTIVEXCONTROL") Then
                    Set objHMIObject = objDoc.HMIObjects.AddActiveXControl(.Cells(lRow, 1).value, .Cells(lRow, 3).value)
                Else
                    Set objHMIObject = objDoc.HMIObjects.AddHMIObject(.Cells(lRow, 1).value, .Cells(lRow, 2).value)
                End If
                objHMIObject.Left = .Cells(lRow, 4).value
                objHMIObject.Top = .Cells(lRow, 5).value
                objHMIObject.Width = .Cells(lRow, 6).value
                objHMIObject.Height = .Cells(lRow, 7).value
                objHMIObject.Layer = .Cells(lRow, 8).value
            End If
  
            Set objHMIObject = Nothing
            lRow = lRow + 1
        Wend
    End With
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



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Document_Activated(CancelForwarding As Boolean)
'VBA76
    MsgBox "The document got the focus." & vbCrLf & "This event (Document_Activated) is raised by the document itself"
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Document_BeforeClose(Cancel As Boolean, CancelForwarding As Boolean)
'VBA77
    MsgBox "Event Document_BeforeClose is raised"
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub objGDApplication_BeforeDocumentClose(ByVal Document As IHMIDocument, Cancel As Boolean)
'VBA78
    MsgBox "The document " & Document.Name & " will be closed after press ok"
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub objGDApplication_BeforeDocumentSave(ByVal Document As IHMIDocument, Cancel As Boolean)
'VBA79
    MsgBox Document.Name & "-saving will start after press ok."
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Document_BeforeHMIObjectDelete(ByVal HMIObject As IHMIObject, Cancel As Boolean, CancelForwarding As Boolean)
'VBA80
    Dim strObjName As String
    Dim strAnswer As String
'
    '"strObjName" contains the name of the deleted object
    strObjName = HMIObject.ObjectName
    strAnswer = MsgBox("Are you sure to delete " & strObjName & "?", vbYesNo)
    If strAnswer = vbNo Then
        'if pressed "No" -> set Cancel to true for prevent delete
        Cancel = True
    End If
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub objGDApplication_BeforeLibraryFolderDelete(ByVal LibObject As HMIFolderItem, Cancel As Boolean)
'VBA81
    MsgBox "The library-folder " & LibObject.Name & " will be delete..."
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub objGDApplication_BeforeLibraryObjectDelete(ByVal LibObject As HMIFolderItem, Cancel As Boolean)
'VBA82
    MsgBox "The object " & LibObject.Name & " will be delete..."
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub objGDApplication_BeforeQuit(Cancel As Boolean)
'VBA83
    MsgBox "The Graphics Designer will be shut down"
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Document_BeforeSave(Cancel As Boolean, CancelForwarding As Boolean)
'VBA84
    MsgBox "The document will be saved..."
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub objGDApplication_DataLanguageChanged(ByVal lCID As Long)
'VBA87
    MsgBox "The datalanguage is changed to " & Application.CurrentDataLanguage & "."
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub objGDApplication_DesktopLanguageChanged(ByVal lCID As Long)
'VBA88
    MsgBox "The desktop-language is changed to " & Application.CurrentDesktopLanguage & "."
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub objGDApplication_DocumentActivated(ByVal Document As IHMIDocument)
'VBA89
    MsgBox "The document " & Document.Name & " got the focus." & vbCrLf & "This event is raised by the application."
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub objGDApplication_DocumentCreated(ByVal Document As IHMIDocument)
'VBA90
    MsgBox Document.Name & " will be created."
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub objGDApplication_DocumentOpened(ByVal Document As IHMIDocument)
'VBA91
    MsgBox Document.Name & " is opened."
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub objGDApplication_DocumentSaved(ByVal Document As IHMIDocument)
'VBA92
    MsgBox Document.Name & " is saved."
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Document_DocumentPropertyChanged(ByVal Property As IHMIProperty, CancelForwarding As Boolean)
'VBA93
    Dim strPropName As String
    '"strPropName" contains the name of the modified property
    strPropName = Property.Name
    MsgBox "The picture-property " & strPropName & " is modified..."
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Document_HMIObjectAdded(ByVal HMIObject As IHMIObject, CancelForwarding As Boolean)
'VBA94
    Dim strObjName As String
'
    '"strObjName" contains the name of the added object
    strObjName = HMIObject.ObjectName
    MsgBox "Object " & strObjName & " is added..."
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Document_HMIObjectMoved(ByVal HMIObject As IHMIObject, CancelForwarding As Boolean)
'VBA95
    Dim strObjName As String
'
    '"strObjName" contains the name of the moved object
    strObjName = HMIObject.ObjectName
    MsgBox "Object " & strObjName & " was moved..."
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Document_HMIObjectPropertyChanged(ByVal Property As IHMIProperty, CancelForwarding As Boolean)
'VBA96
    Dim strObjProp As String
    Dim strObjName As String
    Dim varPropValue As Variant
'
    '"strObjProp" contains the name of the modified property
    '"varPropValue" contains the new value
    strObjProp = Property.Name
    varPropValue = Property.value
'
    '"strObjName" contains the name of the selected object,
    'which property is modified
    strObjName = Property.Application.ActiveDocument.Selection(1).ObjectName
    MsgBox "The property " & strObjProp & " of object " & strObjName & " is modified... " & vbCrLf & "The new value is: " & varPropValue
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Document_HMIObjectResized(ByVal HMIObject As IHMIObject, CancelForwarding As Boolean)
'VBA97
    Dim strObjName As String
'
    '"strObjName" contains the name of the modified object
    strObjName = HMIObject.ObjectName
    MsgBox "The size of " & strObjName & " was modified..."
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub objGDApplication_LibraryFolderRenamed(ByVal LibObject As HMIFolderItem, ByVal OldName As String)
'VBA98
    MsgBox "The Library-folder " & OldName & " is renamed in: " & LibObject.DisplayName
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub objGDApplication_LibraryObjectRenamed(ByVal LibObject As IHMIFolderItem, ByVal OldName As String)
'VBA99
    MsgBox "The object " & OldName & " is renamed in: " & LibObject.DisplayName
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Document_LibraryObjectAdded(ByVal LibObject As IHMIFolderItem, CancelForwarding As Boolean)
'VBA100
    Dim strObjName As String
'
    '"strObjName" contains the name of the added object
    strObjName = LibObject.DisplayName
    MsgBox "Object " & strObjName & " was added to the picture."
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Document_MenuItemClicked(ByVal MenuItem As IHMIMenuItem)
'VBA101
    Dim objMenuItem As HMIMenuItem
    Dim varMenuItemKey As Variant
    Set objMenuItem = MenuItem
'
    '"objMenuItem" contains the clicked menu-item
    '"varMenuItemKey" contains the value of parameter "Key"
    'from the clicked userdefined menu-item
    varMenuItemKey = objMenuItem.Key
    Select Case MenuItemKey
        Case "mItem1_1"
            MsgBox "The first menu-item was clicked!"
    End Select
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub objGDApplication_NewLibraryFolder(ByVal LibObject As IHMIFolderItem)
'VBA102
    MsgBox "The library-folder " & LibObject.DisplayName & " was added."
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub objGDApplication_NewLibraryObject(ByVal LibObject As IHMIFolderItem)
'VBA103
    MsgBox "The object " & LibObject.DisplayName & " was added."
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Document_Opened(CancelForwarding As Boolean)
'VBA104
    MsgBox "The Document is open now..."
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Document_Saved(CancelForwarding As Boolean)
'VBA105
    MsgBox "The document is saved..."
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Document_SelectionChanged(CancelForwarding As Boolean)
'VBA106
    MsgBox "The selection is changed..."
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub objGDApplication_Started()
'VBA107
    MsgBox "The Graphics Designer is started!"
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Document_ToolbarItemClicked(ByVal ToolbarItem As IHMIToolbarItem)
'VBA108
    Dim objToolbarItem As HMIToolbarItem
    Dim varToolbarItemKey As Variant
    Set objToolbarItem = ToolbarItem
'
    '"varToolbarItemKey" contains the value of parameter "Key"
    'from the clicked userdefined toolbar-item
    varToolbarItemKey = objToolbarItem.Key
'
    Select Case varToolbarItemKey
        Case "tItem1_1"
            MsgBox "The first Toolbar-Icon was clicked!"
    End Select
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Document_ViewCreated(ByVal pView As IHMIView, CancelForwarding As Boolean)
'VBA109
    Dim iViewCount As Integer
'
    'To read out the number of views
    iViewCount = pView.Application.ActiveDocument.Views.Count
    MsgBox "A new copy of the picture (number " & iViewCount & ") was created."
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub objGDApplication_WindowStateChanged()
'VBA110
    MsgBox "The state of the application-window is changed!"
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateAndActivateView()
'VBA111
    Dim objView As HMIView
    Set objView = ActiveDocument.Views.Add
    objView.Activate
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeAnalog()
'VBA112
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeAnalog
        .AnalogResultInfos.Add 50, 40
        .AnalogResultInfos.Add 100, 80
        .AnalogResultInfos.ElseCase = 100
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddNewDocument()
'VBA113
    Application.Documents.Add hmiDocumentTypeVisible
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateGroup()
'VBA114
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim objEllipseSegment As HMIEllipseSegment
    Dim objGroup As HMIGroup
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    With objCircle
        .Top = 40
        .Left = 40
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 80
        .Selected = True
    End With
    MsgBox "Objects selected!"
    Set objGroup = ActiveDocument.Selection.CreateGroup
    'Set name for new group-object
    'The name identifies the group-object
    objGroup.ObjectName = "My Group"
    'Add new object to active document...
    Set objEllipseSegment = ActiveDocument.HMIObjects.AddHMIObject("EllipseSegment", "HMIEllipseSegment")
    Set objGroup = ActiveDocument.HMIObjects("My Group")
    '...and add it to the group:
    objGroup.GroupedHMIObjects.Add ("EllipseSegment")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDocumentSpecificCustomToolbar()
'VBA115
    Dim objToolbar As HMIToolbar
    Dim objToolbarItem As HMIToolbarItem
 
    Set objToolbar = ActiveDocument.CustomToolbars.Add("DocToolbar")
    
    'Add toolbar-items to the userdefined toolbar
    Set objToolbarItem = objToolbar.ToolbarItems.InsertToolbarItem(1, "tItem1_1", "Mein erstes Symbol-Icon")
    Set objToolbarItem = objToolbar.ToolbarItems.InsertToolbarItem(3, "tItem1_3", "Mein zweites Symbol-Icon")
'
    'Insert seperatorline between the two tollbaritems
    Set objToolbarItem = objToolbar.ToolbarItems.InsertSeparator(2, "tSeparator1_2")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateViewAndActivateView()
'VBA116
    Dim objView As HMIView
    Set objView = ActiveDocument.Views.Add
    objView.Activate
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateViewAndActivateView()
'VBA117
    Dim objView As HMIView
    Set objView = ActiveDocument.Views.Add
    objView.Activate
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddActionToPropertyTypeVBScript()
'VBA118
    Dim objEvent As HMIEvent
    Dim objVBScript As HMIScriptInfo
    Dim objCircle As HMICircle
    'Create circle in picture. By changing of property "Radius"
    'a VBS-action will be started:
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_AB", "HMICircle")
    Set objEvent = objCircle.Radius.Events(1)
    Set objVBScript = objEvent.Actions.AddAction(hmiActionCreationTypeVBScript)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddActiveXControl()
'VBA119
    Dim objActiveXControl As HMIActiveXControl
    Set objActiveXControl = ActiveDocument.HMIObjects.AddActiveXControl("WinCC_Gauge", "XGAUGE.XGaugeCtrl.1")
    With ActiveDocument
        .HMIObjects("WinCC_Gauge").Top = 40
        .HMIObjects("WinCC_Gauge").Left = 40
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddNewFolderToProjectLibrary()
'VBA120
    Dim objProjectLib As HMISymbolLibrary
    Set objProjectLib = Application.SymbolLibraries(2)
    objProjectLib.FolderItems.AddFolder ("My Folder")
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CopyObjectFromGlobalLibraryToProjectLibrary()
'VBA121
    Dim objGlobalLib As HMISymbolLibrary
    Dim objProjectLib As HMISymbolLibrary
    Set objGlobalLib = Application.SymbolLibraries(1)
    Set objProjectLib = Application.SymbolLibraries(2)
    objProjectLib.FolderItems.AddFolder ("My Folder3")
'
    'copy object from "Global Library" to clipboard
    With objGlobalLib
        .FolderItems(2).Folder.Item(2).Folder.Item(1).CopyToClipboard
    End With
'
    'paste object from clipboard into "Project Library"
    objProjectLib.FolderItems(objProjectLib.FindByDisplayName("My Folder3")).Folder.AddFromClipBoard ("Copy of PC/PLC")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddCircleToActiveDocument()
'VBA122
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("VBA_Circle", "HMICircle")
    objCircle.BackColor = RGB(255, 0, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddObjectFromPictureToProjectLibrary()
'VBA123
    Dim objProjectLib As HMISymbolLibrary
    Dim objCircle As HMICircle
 
    Set objProjectLib = Application.SymbolLibraries(2)
    objProjectLib.FolderItems.AddFolder ("My Folder2")
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle", "HMICircle")
'
    'Add object "Circle" to "Project Library":
    objProjectLib.FolderItems(objProjectLib.FindByDisplayName("My Folder2")).Folder.AddItem "ProjectLib Circle", ActiveDocument.HMIObjects("Circle")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddOLEObjectToActiveDocument()
'VBA124
    Dim objOLEObject As HMIOLEObject
    Set objOLEObject = ActiveDocument.HMIObjects.AddOLEObject("MS Wordpad Document", "Wordpad.Document.1")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AlignSelectedObjectsBottom()
'VBA125
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    With objCircle
        .Top = 40
        .Left = 40
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 80
        .Selected = True
    End With
    MsgBox "Objects selected!"
    ActiveDocument.Selection.AlignBottom
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AlignSelectedObjectsLeft()
'VBA126
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    With objCircle
        .Top = 40
        .Left = 40
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 80
        .Selected = True
    End With
    MsgBox "Objects selected!"
    ActiveDocument.Selection.AlignLeft
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AlignSelectedObjectsRight()
'VBA127
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    With objCircle
        .Top = 40
        .Left = 40
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 80
        .Selected = True
    End With
    MsgBox "Objects selected!"
    ActiveDocument.Selection.AlignRight
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AlignSelectedObjectsTop()
'VBA128
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    With objCircle
        .Top = 40
        .Left = 40
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 80
        .Selected = True
    End With
    MsgBox "Objects selected!"
    ActiveDocument.Selection.AlignTop
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ArrangeMinimizedWindows()
'VBA129
    Application.ArrangeMinimizedWindows
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CascadeWindows()
'VBA130
    Application.CascadeWindows
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CenterSelectedObjectsHorizontally()
'VBA131
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    With objCircle
        .Top = 40
        .Left = 40
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 80
        .Selected = True
    End With
    MsgBox "Objects selected!"
    ActiveDocument.Selection.CenterHorizontally
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CenterSelectedObjectsVertically()
'VBA132
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    With objCircle
        .Top = 40
        .Left = 40
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 80
        .Selected = True
    End With
    MsgBox "Objects selected!"
    ActiveDocument.Selection.CenterVertically
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CloseDocumentUsingTheFileName()
'VBA134
    Dim strFile As String
    strFile = Application.ApplicationDataPath & "test.pdl"
    Application.Documents.Close (strFile)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CloseDocumentUsingActiveDocument()
'VBA135
    ActiveDocument.Close
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CloseAllDocuments()
'VBA136
    Application.Documents.CloseAll
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ConvertDynamicDialogToScript()
'VBA137
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
'
    'Create dynamic
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
'
    'configure dynamic. "ResultType" defines the valuerange-type:
    With objDynDialog
        .ResultType = hmiResultTypeAnalog
        .AnalogResultInfos.Add 50, 40
        .AnalogResultInfos.Add 100, 80
        .AnalogResultInfos.ElseCase = 100
        MsgBox "The dynamic-dialog will be changed into a C-script."
        .ConvertToScript
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CopySelectionToNewDocument()
'VBA138
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim iNewDoc As Integer
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    With objCircle
        .Top = 40
        .Left = 40
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 80
        .Selected = True
    End With
    MsgBox "Objects selected!"
    'Instead of "ActiveDocument.CopySelection" you can also write:
    '"ActiveDocument.Selection.CopySelection".
    ActiveDocument.CopySelection
    Application.Documents.Add hmiDocumentTypeVisible
    iNewDoc = Application.Documents.Count
    Application.Documents(iNewDoc).PasteClipboard
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CopyObjectFromGlobalLibraryToProjectLibrary()
'VBA139
    Dim objGlobalLib As HMISymbolLibrary
    Dim objProjectLib As HMISymbolLibrary
    Set objGlobalLib = Application.SymbolLibraries(1)
    Set objProjectLib = Application.SymbolLibraries(2)
    objProjectLib.FolderItems.AddFolder ("My Folder3")
'
    'copy object from "Global Library" to clipboard
    With objGlobalLib
        .FolderItems(2).Folder.Item(2).Folder.Item(1).CopyToClipboard
    End With
'
    'paste object from clipboard into "Project Library"
    objProjectLib.FolderItems(objProjectLib.FindByDisplayName("My Folder3")).Folder.AddFromClipBoard ("Copy of PC/PLC")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateCustomizedObject()
'VBA140
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim objCustObject As HMICustomizedObject
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    With objCircle
        .Top = 40
        .Left = 40
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 80
        .Selected = True
    End With
    MsgBox "Objects selected!"
    Set objCustObject = ActiveDocument.Selection.CreateCustomizedObject
    objCustObject.ObjectName = "myCustomizedObject"
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicAsVariableDirectToProperty()
'VBA141
    Dim objVariableTrigger As HMIVariableTrigger
    Dim objCircle As HMICircle
 
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("MyCircle", "HMICircle")
    'Make property "Top" dynamic:
    Set objVariableTrigger = objCircle.Top.CreateDynamic(hmiDynamicCreationTypeVariableDirect, "NewDynamic")
'
    'Define cycle-time
    With objVariableTrigger
        .CycleType = hmiCycleType_2s
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateGroup()
'VBA142
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim objGroup As HMIGroup
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    With objCircle
        .Top = 40
        .Left = 40
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 80
        .Selected = True
    End With
    MsgBox "Objects selected!"
    Set objGroup = ActiveDocument.Selection.CreateGroup
    objGroup.ObjectName = "myGroup"
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ObjectDelete()
'VBA143
    ActiveDocument.HMIObjects(1).Delete
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DeleteActionOfCircleAB()
'VBA144
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects("Circle_AB")
    objCircle.Radius.Events(1).Actions(1).Delete
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DeleteAllSelectedObjects()
'VBA145
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    With objCircle
        .Top = 40
        .Left = 40
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 80
        .Selected = True
    End With
    MsgBox "Objects selected!"
    ActiveDocument.Selection.DeleteAll
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DeleteDynamicFromObjectMeinKreis()
'VBA146
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects("MyCircle")
    objCircle.Top.DeleteDynamic
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SelectObjectsAndDeselectThemAgain()
'VBA147
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    With objCircle
        .Top = 40
        .Left = 40
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 80
        .Selected = True
    End With
    MsgBox "Objects created and selected!"
    ActiveDocument.Selection.DeselectAll
    MsgBox "Objects deselected!"
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DuplicateSelectedObjects()
'VBA149
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    With objCircle
        .Top = 40
        .Left = 40
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 80
        .Selected = True
    End With
    MsgBox "Objects created and selected!"
    ActiveDocument.Selection.DuplicateSelection
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EvenlySpaceObjectsHorizontally()
'VBA150
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim objEllipse As HMIEllipse
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    Set objEllipse = ActiveDocument.HMIObjects.AddHMIObject("sEllipse", "HMIEllipse")
    With objCircle
        .Top = 30
        .Left = 0
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 42
        .Selected = True
    End With
    With objEllipse
        .Top = 48
        .Left = 162
        .BackColor = RGB(255, 0, 0)
        .Selected = True
    End With
    MsgBox "Objects created and selected!"
    ActiveDocument.Selection.EvenlySpaceHorizontally
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EvenlySpaceObjectsVertically()
'VBA151
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim objEllipse As HMIEllipse
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    Set objEllipse = ActiveDocument.HMIObjects.AddHMIObject("sEllipse", "HMIEllipse")
    With objCircle
        .Top = 30
        .Left = 0
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 42
        .Selected = True
    End With
    With objEllipse
        .Top = 48
        .Left = 162
        .BackColor = RGB(255, 0, 0)
        .Selected = True
    End With
    MsgBox "Objects created and selected"
    ActiveDocument.Selection.EvenlySpaceVertically
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub FindObjectsByType()
'VBA153
    Dim colSearchResults As HMICollection
    Dim objMember As HMIObject
    Dim iResult As Integer
    Dim strName As String
    Set colSearchResults = ActiveDocument.HMIObjects.Find(ObjectType:="HMICircle")
    For Each objMember In colSearchResults
        iResult = colSearchResults.Count
        strName = objMember.ObjectName
        MsgBox "Found: " & CStr(iResult) & vbCrLf & "objectname: " & strName)
    Next objMember
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub FindObjectInSymbolLibrary()
'VBA154
    Dim objGlobalLib As HMISymbolLibrary
    Dim objFItem As HMIFolderItem
    Set objGlobalLib = Application.SymbolLibraries(1)
    Set objFItem = objGlobalLib.FindByDisplayName("PC")
    MsgBox objFItem.DisplayName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub FlipObjectHorizontally()
'VBA155
    Dim objStaticText As HMIStaticText
    Dim strPropertyName As String
    Dim iPropertyValue As Integer
    Set objStaticText = ActiveDocument.HMIObjects.AddHMIObject("Textfield", "HMIStaticText")
    strPropertyName = objStaticText.Properties("Text").Name
    With objStaticText
        .Width = 120
        .Text = "Sample Text"
        .Selected = True
        iPropertyValue = .AlignmentTop
        MsgBox "Value of '" & strPropertyName & "' before flip: " & iPropertyValue
        ActiveDocument.Selection.FlipHorizontally
        iPropertyValue = objStaticText.AlignmentTop
        MsgBox "Value of '" & strPropertyName & "' after flip: " & iPropertyValue
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub FlipObjectVertically()
'VBA156
    Dim objStaticText As HMIStaticText
    Dim strPropertyName As String
    Dim iPropertyValue As Integer
    Set objStaticText = ActiveDocument.HMIObjects.AddHMIObject("Textfield", "HMIStaticText")
    strPropertyName = objStaticText.Properties("Text").Name
    With objStaticText
        .Width = 120
        .Text = "Sample Text"
        .Selected = True
        .AlignmentLeft = 0
        iPropertyValue = .AlignmentLeft
        MsgBox "Value of '" & strPropertyName & "' before flip: " & iPropertyValue
        ActiveDocument.Selection.FlipVertically
        iPropertyValue = objStaticText.AlignmentLeft
        MsgBox "Value of '" & strPropertyName & "' after flip: " & iPropertyValue
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowDisplayName()
'VBA157
    Dim objGlobalLib As HMISymbolLibrary
    Dim objFItem As HMIFolderItem
    Set objGlobalLib = Application.SymbolLibraries(1)
    Set objFItem = objGlobalLib.GetItemByPath("\Folder1\Folder2\Object1")
    MsgBox objFItem.DisplayName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ToolbarItem_InsertFromMenuItem()
'VBA158
    Dim objMenu As HMIMenu
    Dim objToolbarItem As HMIToolbarItem
    Dim objToolbar As HMIToolbar
    Dim objMenuItem As HMIMenuItem
    Set objMenu = Application.CustomMenus.InsertMenu(1, "Menu1", "TestMenu")
'
    '*************************************************
    '* Note:
    '* The object-reference has to be unique.
    '*************************************************
'
    Set objMenuItem = Application.CustomMenus(1).MenuItems.InsertMenuItem(1, "MenuItem1", "Hello World")
    Application.CustomMenus(1).MenuItems(1).Macro = "HelloWorld"
    Set objToolbar = Application.CustomToolbars.Add("Toolbar1")
    Set objToolbarItem = Application.CustomToolbars(1).ToolbarItems.InsertFromMenuItem(1, "ToolbarItem1", objMenuItem, "Call's Hello World of TestMenu")
End Sub

Sub HelloWorld()
    MsgBox "Procedure 'HelloWorld()' is execute."
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateDocumentMenus()
'VBA159
    Dim objDocMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
    Dim objSubMenu As HMIMenuItem
'
    Set objDocMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DocMenu1", "Doc_Menu_1")
'
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(1, "dmItem1_1", "First MenuItem")
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(2, "dmItem1_2", "Second MenuItem")
'
    'Insert a dividing rule into custumized menu:
    Set objMenuItem = objDocMenu.MenuItems.InsertSeparator(3, "dSeparator1_3")
'
    Set objSubMenu = objDocMenu.MenuItems.InsertSubMenu(4, "dSubMenu1_4", "First SubMenu")
'
    Set objMenuItem = objSubMenu.SubMenu.InsertMenuItem(5, "dmItem1_5", "First item in sub-menu")
    Set objMenuItem = objSubMenu.SubMenu.InsertMenuItem(6, "dmItem1_6", "Second item in sub-menu")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateDocumentMenus()
'VBA160
    Dim objDocMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
    Dim objSubMenu As HMIMenuItem
'
    Set objDocMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DocMenu1", "Doc_Menu_1")
'
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(1, "dmItem1_1", "First MenuItem")
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(2, "dmItem1_2", "Second MenuItem")
'
    'Insert a dividing rule to customized menu:
    Set objMenuItem = objDocMenu.MenuItems.InsertSeparator(3, "dSeparator1_3")
'
    Set objSubMenu = objDocMenu.MenuItems.InsertSubMenu(4, "dSubMenu1_4", "First SubMenu")
'
    Set objMenuItem = objSubMenu.SubMenu.InsertMenuItem(5, "dmItem1_5", "First item in sub-menu")
    Set objMenuItem = objSubMenu.SubMenu.InsertMenuItem(6, "dmItem1_6", "Second item in sub-menu")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDocumentSpecificCustomToolbar()
'VBA161
    Dim objToolbar As HMIToolbar
    Dim objToolbarItem As HMIToolbarItem
 
    Set objToolbar = ActiveDocument.CustomToolbars.Add("DocToolbar")
    
    'Add toolbar-item to userdefined toolbar
    Set objToolbarItem = objToolbar.ToolbarItems.InsertToolbarItem(1, "tItem1_1", "First symbol-icon")
    Set objToolbarItem = objToolbar.ToolbarItems.InsertToolbarItem(3, "tItem1_3", "Second symbol-icon")
'
    'Insert dividing rule between first and second symbol-icon
    Set objToolbarItem = objToolbar.ToolbarItems.InsertSeparator(2, "tSeparator1_2")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateDocumentMenus()
'VBA162
    Dim objDocMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
    Dim objSubMenu As HMIMenuItem
'
    Set objDocMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DocMenu1", "Doc_Menu_1")
'
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(1, "dmItem1_1", "First MenuItem")
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(2, "dmItem1_2", "Second MenuItem")
'
    'Insert a dividing rule to customized menu:
    Set objMenuItem = objDocMenu.MenuItems.InsertSeparator(3, "dSeparator1_3")
'
    Set objSubMenu = objDocMenu.MenuItems.InsertSubMenu(4, "dSubMenu1_4", "First SubMenu")
'
    Set objMenuItem = objSubMenu.SubMenu.InsertMenuItem(5, "dmItem1_5", "First item in sub-menu")
    Set objMenuItem = objSubMenu.SubMenu.InsertMenuItem(6, "dmItem1_6", "Second item in sub-menu")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDocumentSpecificCustomToolbar()
'VBA163
    Dim objToolbar As HMIToolbar
    Dim objToolbarItem As HMIToolbarItem
 
    Set objToolbar = ActiveDocument.CustomToolbars.Add("DocToolbar")
    
    'Add toolbar-item to userdefined toolbar
    Set objToolbarItem = objToolbar.ToolbarItems.InsertToolbarItem(1, "tItem1_1", "First symbol-icon")
    Set objToolbarItem = objToolbar.ToolbarItems.InsertToolbarItem(3, "tItem1_3", "Second symbol-icon")
'
    'Insert dividing rule between first and second symbol-icon
    Set objToolbarItem = objToolbar.ToolbarItems.InsertSeparator(2, "tSeparator1_2")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub IsCSLayerVisible()
'VBA164
    Dim objView As HMIView
    Dim strLayerName As String
    Dim iLayerIdx As Integer
    Set objView = ActiveDocument.Views(1)
    objView.Activate
    iLayerIdx = 2
    strLayerName = ActiveDocument.Layers(iLayerIdx).Name
    If objView.IsCSLayerVisible(iLayerIdx) = True Then
        MsgBox "CS " & strLayerName & " is visible"
    Else
        MsgBox "CS " & strLayerName & " is invisible"
    End If
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RTLayerVisibility()
'VBA165
    Dim strLayerName As String
    Dim iLayerIdx As Integer
    iLayerIdx = 2
    strLayerName = ActiveDocument.Layers(iLayerIdx).Name
    If ActiveDocument.IsRTLayerVisible(iLayerIdx) = True Then
        MsgBox "RT " & strLayerName & " is visible"
    Else
        MsgBox "RT " & strLayerName & " is invisible"
    End If
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowDocumentNameLongVersion()
'VBA166
    Dim strDocName As String
    strDocName = Application.Documents.Item(3).Name
    MsgBox strDocName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowDocumentNameShortVersion()
'VBA167
    Dim strDocName As String
    strDocName = Application.Documents(3).Name
    MsgBox strDocName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ExampleForLanguageFonts()
'VBA168
    Dim objLangFonts As HMILanguageFonts
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    objButton.Text = "Hello"
    Set objLangFonts = objButton.LDFonts
'
    'To make fontsettings for french:
    With objLangFonts.ItemByLCID(1036)
        .Family = "Courier New"
        .Bold = True
        .Italic = False
        .Underlined = True
        .Size = 12
    End With
'
    'To make fontsettings for english:
    With objLangFonts.ItemByLCID(1033)
        .Family = "Times New Roman"
        .Bold = False
        .Italic = True
        .Underlined = False
        .Size = 14
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub LoadDefaultConfig()
'VBA169
    Application.LoadDefaultConfig ("Test.PDD")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub MoveSelectionToNewPostion()
'VBA172
    Dim nPosX As Long
    Dim nPosY As Long
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    With objCircle
        .Top = 40
        .Left = 40
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 80
        .Selected = True
    End With
    MsgBox "Objects selected!"
    nPosX = 30
    nPosY = 40
    'Instead of next you can write "ActiveDocument.Selection.MoveSelection".
    ActiveDocument.MoveSelection nPosX, nPosY
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub MoveObjectOneLevelBackward()
'VBA173
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    With objCircle
        .Top = 40
        .Left = 40
        .Selected = False
    End With
    With objRectangle
        .Top = 40
        .Left = 40
        .Width = 100
        .Height = 50
        .BackColor = RGB(255, 0, 255)
        .Selected = True
    End With
    MsgBox "Objects created and selected!"
    ActiveDocument.Selection.BackwardOneLevel
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub MoveObjectOneLevelForward()
'VBA174
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    With objCircle
        .Top = 40
        .Left = 40
        .Selected = True
    End With
    With objRectangle
        .Top = 40
        .Left = 40
        .Width = 100
        .Height = 50
        .BackColor = RGB(255, 0, 255)
        .Selected = False
    End With
    MsgBox "Objects created and selected!"
    ActiveDocument.Selection.ForwardOneLevel
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub OpenDocument()
'VBA175
    Application.Documents.Open ("Test.PDL", hmiDocumentTypeVisible)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CopySelectionToNewDocument()
'VBA176
    Dim iNewDoc As String
    ActiveDocument.CopySelection
    Application.Documents.Add hmiDocumentTypeVisible
    iNewDoc = Application.Documents.Count
    Application.Documents(iNewDoc).PasteClipboard
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateAndPrintView()
'VBA177
    Dim objView As HMIView
    Set objView = ActiveDocument.Views.Add
    objView.Activate
    objView.PrintDocument
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ToPrintProjectDocumentation()
'VBA178
    ActiveDocument.PrintProjectDocumentation
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RemoveObjectFromGroup()
'VBA179
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim objEllipse As HMIEllipse
    Dim objGroup As HMIGroup
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    Set objEllipse = ActiveDocument.HMIObjects.AddHMIObject("sEllipse", "HMIEllipse")
    With objCircle
        .Top = 30
        .Left = 0
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 42
        .Selected = True
    End With
    With objEllipse
        .Top = 48
        .Left = 162
        .Width = 40
        .BackColor = RGB(255, 0, 0)
        .Selected = True
    End With
    MsgBox "Objects selected!"
    Set objGroup = ActiveDocument.Selection.CreateGroup
    MsgBox "Group-object is created."
    objGroup.GroupedHMIObjects.Remove ("sEllipse")
    MsgBox "The ellipse is removed from group-object."
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RotateGroupObject()
'VBA180
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim objGroup As HMIGroup
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    With objRectangle
        .Top = 30
        .Left = 30
        .Width = 80
        .Height = 40
        .Selected = True
    End With
    With objCircle
        .Top = 30
        .Left = 30
        .BackColor = RGB(255, 255, 255)
        .Selected = True
    End With
    MsgBox "Objects selected!"
    Set objGroup = ActiveDocument.Selection.CreateGroup
    MsgBox "Group-object created."
    objGroup.Selected = True
    ActiveDocument.Selection.Rotate
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ApplySameHeightToSelectedObjects()
'VBA181
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim objEllipse As HMIEllipse
    Dim objGroup As HMIGroup
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    Set objEllipse = ActiveDocument.HMIObjects.AddHMIObject("sEllipse", "HMIEllipse")
    With objCircle
        .Top = 30
        .Left = 0
        .Height = 15
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 42
        .Height = 40
        .Selected = True
    End With
    With objEllipse
        .Top = 48
        .Left = 162
        .Width = 40
        .Height = 120
        .BackColor = RGB(255, 0, 0)
        .Selected = True
    End With
    MsgBox "Objects selected!"
    ActiveDocument.Selection.SameHeight
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ApplySameWidthToSelectedObjects()
'VBA182
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim objEllipse As HMIEllipse
    Dim objGroup As HMIGroup
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    Set objEllipse = ActiveDocument.HMIObjects.AddHMIObject("sEllipse", "HMIEllipse")
    With objCircle
        .Top = 30
        .Left = 0
        .Width = 15
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 42
        .Width = 40
        .Selected = True
    End With
    With objEllipse
        .Top = 48
        .Left = 162
        .Width = 120
        .BackColor = RGB(255, 0, 0)
        .Selected = True
    End With
    MsgBox "Objects selected!"
    ActiveDocument.Selection.SameWidth
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ApplySameWidthAndHeightToSelectedObjects()
'VBA183
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim objEllipse As HMIEllipse
    Dim objGroup As HMIGroup
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    Set objEllipse = ActiveDocument.HMIObjects.AddHMIObject("sEllipse", "HMIEllipse")
    With objCircle
        .Top = 30
        .Left = 0
        .Height = 15
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 42
        .Width = 25
        .Height = 40
        .Selected = True
    End With
    With objEllipse
        .Top = 48
        .Left = 162
        .Width = 40
        .Height = 120
        .BackColor = RGB(255, 0, 0)
        .Selected = True
    End With
    MsgBox "Objects selected!"
    ActiveDocument.Selection.SameWidthAndHeight
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SaveDocument()
'VBA184
    ActiveDocument.Save
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SaveAllDocuments()
'VBA185
    Application.Documents.SaveAll
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SaveDocumentAs()
'VBA186
    ActiveDocument.SaveAs ("Test2.PDL")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SaveDefaultConfig()
'VBA187
    Application.SaveDefaultConfig ("Test.PDD")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SelectAllObjectsInActiveDocument()
'VBA188
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim objEllipse As HMIEllipse
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    Set objEllipse = ActiveDocument.HMIObjects.AddHMIObject("sEllipse", "HMIEllipse")
    With objCircle
        .Top = 30
        .Left = 0
        .Height = 15
    End With
    With objRectangle
        .Top = 80
        .Left = 42
        .Width = 25
        .Height = 40
    End With
    With objEllipse
        .Top = 48
        .Left = 162
        .Width = 40
        .Height = 120
        .BackColor = RGB(255, 0, 0)
    End With
    ActiveDocument.Selection.SelectAll
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SetCSLayerVisible()
'VBA189
    Dim objView As HMIView
    Set objView = ActiveDocument.Views.Add
    objView.Activate
    objView.SetCSLayerVisible 2, False
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ConfigureSettingsOfLayer()
'VBA190
    Dim objLayer As HMILayer
    Set objLayer = ActiveDocument.Layers(1)
    With objLayer
        'Configure "Layer 0"
        .MinZoom = 10
        .MaxZoom = 100
        .Name = "Configured with VBA"
    End With
    'Define decluttering of objects:
    With ActiveDocument
        .LayerDecluttering = True
        .ObjectSizeDecluttering = True
        .SetDeclutterObjectSize 50, 100
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SetRTLayerVisibleWithVBA()
'VBA191
    ActiveDocument.SetRTLayerVisible 1, False
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowPropertiesDialog()
'VBA192
    Application.ShowPropertiesDialog
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowSymbolLibraryDialog()
'VBA193
    Application.ShowSymbolLibraryDialog
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowTagDialog()
'VBA194
    Application.ShowTagDialog
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub TileWindowsHorizontally()
'VBA195
    Application.TileWindowsHorizontally
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub TileWindowsVertically()
'VBA196
    Application.TileWindowsVertically
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SendObjectToBack()
'VBA197
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    With objCircle
        .Top = 40
        .Left = 40
        .Selected = False
    End With
    With objRectangle
        .Top = 40
        .Left = 40
        .Width = 100
        .Height = 50
        .BackColor = RGB(255, 0, 255)
        .Selected = True
    End With
    MsgBox "The objects circle and rectangle are created" & vbCrLf & "Only the rectangle is selected!"
    ActiveDocument.Selection.SendToBack
    MsgBox "The selection is moved to the back."
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub MoveObjectToFront()
'VBA198
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    With objCircle
        .Top = 40
        .Left = 40
        .Selected = True
    End With
    With objRectangle
        .Top = 40
        .Left = 40
        .Width = 100
        .Height = 50
        .BackColor = RGB(255, 0, 255)
        .Selected = False
    End With
    MsgBox "The objects circle and rectangle are created" & vbCrLf & "Only the circle is selected!"
    ActiveDocument.Selection.BringToFront
    MsgBox "The selection is moved to the front."
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DissolveGroup()
'VBA199
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim objEllipse As HMIEllipse
    Dim objGroup As HMIGroup
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    Set objEllipse = ActiveDocument.HMIObjects.AddHMIObject("sEllipse", "HMIEllipse")
    With objCircle
        .Top = 30
        .Left = 0
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 42
        .Selected = True
    End With
    With objEllipse
        .Top = 48
        .Left = 162
        .Width = 40
        .BackColor = RGB(255, 0, 0)
        .Selected = True
    End With
    MsgBox "Objects selected!"
    Set objGroup = ActiveDocument.Selection.CreateGroup
    MsgBox "Group-object is created."
    With objGroup
        .Left = 120
        .Top = 300
        MsgBox "Group-object is moved."
        .UnGroup
        MsgBox "Group is dissolved."
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub Add3DBarGraph()
'VBA200
    Dim obj3DBarGraph As HMI3DBarGraph
    Set obj3DBarGraph = ActiveDocument.HMIObjects.AddHMIObject("3DBar", "HMI3DBarGraph")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub Edit3DBarGraph()
'VBA201
    Dim obj3DBarGraph As HMI3DBarGraph
    Set obj3DBarGraph = ActiveDocument.HMIObjects("3DBar")
    obj3DBarGraph.BorderColor = RGB(255, 0, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNameOfFirstSelectedObject()
'VBA202
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name from the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateVBActionToClickedEvent()
'VBA203
    Dim objButton As HMIButton
    Dim objCircle As HMICircle
    Dim objVBScript As HMIScriptInfo
    Dim strVBCode As String
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_VB", "HMICircle")
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    With objCircle
        .Top = 100
        .Left = 100
        .BackColor = RGB(255, 0, 0)
    End With
    With objButton
        .Top = 10
        .Left = 10
        .Text = "Increase Radius"
    End With
    'define event and assign sourcecode to it:
    Set objVBScript = objButton.Events(1).Actions.AddAction(hmiActionCreationTypeVBScript)
    strVBCode = "Dim myCircle" & vbCrLf & "Set myCircle = "
    strVBCode = strVBCode & "HMIRuntime.ActiveScreen.ScreenItems(""Circle_VB"")"
    strVBCode = strVBCode & vbCrLf & "myCircle.Radius = myCircle.Radius + 5"
    With objVBScript
        .SourceCode = strVBCode
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddActiveXControl()
'VBA204
    Dim objActiveXControl As HMIActiveXControl
    Set objActiveXControl = ActiveDocument.HMIObjects.AddActiveXControl("WinCC_Gauge", "XGAUGE.XGaugeCtrl.1")
    With ActiveDocument
        .HMIObjects("WinCC_Gauge").Top = 40
        .HMIObjects("WinCC_Gauge").Left = 40
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeAnalog()
'VBA206
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeAnalog
        .AnalogResultInfos.Add 50, 40
        .AnalogResultInfos.Add 100, 80
        .AnalogResultInfos.ElseCase = 100
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowAnalogResultInfosOfCircleRadius()
'VBA207
    Dim colAResultInfos As HMIAnalogResultInfos
    Dim objAResultInfo As HMIAnalogResultInfo
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Dim iAnswer As Integer
    Dim varRange As Variant
    Dim varValue As Variant
    Set objCircle = ActiveDocument.HMIObjects("Circle_A")
    Set objDynDialog = objCircle.Radius.Dynamic
    Set colAResultInfos = objDynDialog.AnalogResultInfos
    For Each objAResultInfo In colAResultInfos
        varRange = objAResultInfo.RangeTo
        varValue = objAResultInfo.value
        iAnswer = MsgBox("Ranges of values from Circle_A-Radius:" & vbCrLf & "Range of value to: " & varRange & vbCrLf & "Value of property: " & varValue, vbOKCancel)
        If vbCancel = iAnswer Then Exit For
    Next objAResultInfo
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowApplicationVersion()
'VBA208
    MsgBox Application.Version
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddApplicationWindow()
'VBA209
    Dim objApplicationWindow As HMIApplicationWindow
    Set objApplicationWindow = ActiveDocument.HMIObjects.AddHMIObject("AppWindow", "HMIApplicationWindow")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditApplicationWindow()
'VBA210
    Dim objApplicationWindow As HMIApplicationWindow
    Set objApplicationWindow = ActiveDocument.HMIObjects("AppWindow")
    objApplicationWindow.Sizeable = True
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNameOfFirstSelectedObject()
'VBA211
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name from the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddBarGraph()
'VBA212
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditBarGraph()
'VBA213
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects("Bar1")
    objBarGraph.BorderColor = RGB(255, 0, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNameOfFirstSelectedObject()
'VBA214
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name from the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeBinary()
'VBA215
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_C", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeBool
        .BinaryResultInfo.NegativeValue = 20
        .BinaryResultInfo.PositiveValue = 40
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeBit()
'VBA216
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_B", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeBit
        .BitResultInfo.BitNumber = 1
        .BitResultInfo.BitSetValue = 40
        .BitResultInfo.BitNotSetValue = 80
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddButton()
'VBA217
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button", "HMIButton")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditButton()
'VBA218
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects("Button")
    objButton.BorderColor = RGB(255, 0, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNameOfFirstSelectedObject()
'VBA219
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name from the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddCheckBox()
'VBA220
    Dim objCheckBox As HMICheckBox
    Set objCheckBox = ActiveDocument.HMIObjects.AddHMIObject("CheckBox", "HMICheckBox")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditCheckBox()
'VBA221
    Dim objCheckBox As HMICheckBox
    Set objCheckBox = ActiveDocument.HMIObjects("CheckBox")
    objCheckBox.BorderColor = RGB(255, 0, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNameOfFirstSelectedObject()
'VBA222
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name from the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddCircle()
'VBA223
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle", "HMICircle")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditCircle()
'VBA224
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects("Circle")
    objCircle.BorderColor = RGB(255, 0, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNameOfFirstSelectedObject()
'VBA225
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name from the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddCiruclarArc()
'VBA226
    Dim objCiruclarArc As HMICircularArc
    Set objCiruclarArc = ActiveDocument.HMIObjects.AddHMIObject("CircularArc", "HMICircularArc")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditCiruclarArc()
'VBA227
    Dim objCiruclarArc As HMICircularArc
    Set objCiruclarArc = ActiveDocument.HMIObjects("CircularArc")
    objCiruclarArc.BorderColor = RGB(255, 0, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNameOfFirstSelectedObject()
'VBA228
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name from the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CountConnectionPoints()
'VBA229
    Dim objRectangle As HMIRectangle
    Dim objConnPoints As HMIConnectionPoints
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    Set objConnPoints = ActiveDocument.HMIObjects("Rectangle1").ConnectionPoints
    MsgBox "Rectangle1 has " & objConnPoints.Count & " connectionpoints."
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateCustomizedObject()
'VBA230
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim objCustomizedObject As HMICustomizedObject
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle_B", "HMIRectangle")
    With objCircle
        .Left = 10
        .Top = 10
        .Selected = True
    End With
    With objRectangle
        .Left = 50
        .Top = 50
        .Selected = True
    End With
    MsgBox "objects created and selected!"
    Set objCustomizedObject = ActiveDocument.Selection.CreateCustomizedObject
    objCustomizedObject.ObjectName = "Customer-Object"
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditCustomizedObject()
'VBA231
    Dim objCustomizedObject As HMICustomizedObject
    Set objCustomizedObject = ActiveDocument.HMIObjects("Customer-Object")
    MsgBox objCustomizedObject.ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowFirstObjectOfCollection()
'VBA232
    Dim strName As String
    strName = ActiveDocument.Application.AvailableDataLanguages(1).LanguageName
    MsgBox strName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowDataLanguage()
'VBA233
    Dim colDataLanguages As HMIDataLanguages
    Dim objDataLanguage As HMIDataLanguage
    Dim strLanguages As String
    Dim iCount As Integer
    iCount = 0
    Set colDataLanguages = Application.AvailableDataLanguages
    For Each objDataLanguage In colDataLanguages
        If "" <> strLanguages Then strLanguages = strLanguages & "/"
        strLanguages = strLanguages & objDataLanguage.LanguageName & " "
        'Every 15 items of datalanguages output in a messagebox
        If 0 = iCount Mod 15 And 0 <> iCount Then
            MsgBox strLanguages
            strLanguages = ""
        End If
        iCount = iCount + 1
    Next objDataLanguage
    MsgBox strLanguages
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DirectConnection()
'VBA234
    Dim objButton As HMIButton
    Dim objRectangleA As HMIRectangle
    Dim objRectangleB As HMIRectangle
    Dim objEvent As HMIEvent
    Dim objDirConnection As HMIDirectConnection
    Set objRectangleA = ActiveDocument.HMIObjects.AddHMIObject("Rectangle_A", "HMIRectangle")
    Set objRectangleB = ActiveDocument.HMIObjects.AddHMIObject("Rectangle_B", "HMIRectangle")
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    With objRectangleA
        .Top = 100
        .Left = 100
    End With
    With objRectangleB
        .Top = 250
        .Left = 400
        .BackColor = RGB(255, 0, 0)
    End With
    With objButton
        .Top = 10
        .Left = 10
        .Width = 90
        .Height = 50
        .Text = "SetPosition"
    End With
'
    'Directconnection is initiated on mouseclick:
    Set objDirConnection = objButton.Events(1).Actions.AddAction(hmiActionCreationTypeDirectConnection)
    With objDirConnection
        'Sourceobject: Property "Top" of "Rectangle_A"
        .SourceLink.Type = hmiSourceTypeProperty
        .SourceLink.ObjectName = "Rectangle_A"
        .SourceLink.AutomationName = "Top"
'
        'Targetobject: Property "Left" of "Rectangle_B"
        .DestinationLink.Type = hmiDestTypeProperty
        .DestinationLink.ObjectName = "Rectangle_B"
        .DestinationLink.AutomationName = "Left"
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DirectConnection()
'VBA235
    Dim objButton As HMIButton
    Dim objRectangleA As HMIRectangle
    Dim objRectangleB As HMIRectangle
    Dim objEvent As HMIEvent
    Dim objDirConnection As HMIDirectConnection
    Set objRectangleA = ActiveDocument.HMIObjects.AddHMIObject("Rectangle_A", "HMIRectangle")
    Set objRectangleB = ActiveDocument.HMIObjects.AddHMIObject("Rectangle_B", "HMIRectangle")
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    With objRectangleA
        .Top = 100
        .Left = 100
    End With
    With objRectangleB
        .Top = 250
        .Left = 400
        .BackColor = RGB(255, 0, 0)
    End With
    With objButton
        .Top = 10
        .Left = 10
        .Width = 90
        .Height = 50
        .Text = "SetPosition"
    End With
'
    'Directconnection is initiated on mouseclick:
    Set objDirConnection = objButton.Events(1).Actions.AddAction(hmiActionCreationTypeDirectConnection)
    With objDirConnection
        'Sourceobject: Property "Top" of "Rectangle_A"
        .SourceLink.Type = hmiSourceTypeProperty
        .SourceLink.ObjectName = "Rectangle_A"
        .SourceLink.AutomationName = "Top"
'
        'Targetobject: Property "Left" of "Rectangle_B"
        .DestinationLink.Type = hmiDestTypeProperty
        .DestinationLink.ObjectName = "Rectangle_B"
        .DestinationLink.AutomationName = "Left"
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowFirstObjectOfCollection()
'VBA236
    Dim strName As String
    strName = Application.Documents(3).Name
    MsgBox strName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SaveDocumentAs()
'VBA237
    Application.Documents(3).SaveAs ("CopyOfPicture1")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowDocuments()
'VBA238
    Dim colDocuments As Documents
    Dim objDocument As Document
    Set colDocuments = Application.Documents
    For Each objDocument In colDocuments
        MsgBox objDocument.Name
    Next objDocument
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddNewDocument()
'VBA239
    Dim objDocument As Document
    Set objDocument = Application.Documents.Add hmiDocumentTypeVisible
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeAnalog()
'VBA240
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeAnalog
        .Trigger.VariableTriggers.Add "NewDynamic1", hmiCycleType_5s
        .AnalogResultInfos.Add 50, 40
        .AnalogResultInfos.Add 100, 80
        .AnalogResultInfos.ElseCase = 100
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddEllipse()
'VBA241
    Dim objEllipse As HMIEllipse
    Set objEllipse = ActiveDocument.HMIObjects.AddHMIObject("Ellipse", "HMIEllipse")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditEllipse()
'VBA242
    Dim objEllipse As HMIEllipse
    Set objEllipse = ActiveDocument.HMIObjects("Ellipse")
    objEllipse.BorderColor = RGB(255, 0, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNameOfFirstSelectedObject()
'VBA243
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name from the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddEllipseArc()
'VBA244
    Dim objEllipseArc As HMIEllipseArc
    Set objEllipseArc = ActiveDocument.HMIObjects.AddHMIObject("EllipseArc", "HMIEllipseArc")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditEllipseArc()
'VBA245
    Dim objEllipseArc As HMIEllipseArc
    Set objEllipseArc = ActiveDocument.HMIObjects("EllipseArc")
    objEllipseArc.BorderColor = RGB(255, 0, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNameOfFirstSelectedObject()
'VBA246
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name from the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditDefaultPropertiesOfEllipseArc()
'VBA247
    Dim objEllipseArc As HMIEllipseArc
    Set objEllipseArc = Application.DefaultHMIObjects("HMIEllipseArc")
    objEllipseArc.BorderColor = RGB(255, 255, 0)
    'create new "EllipseArc"-object
    Set objEllipseArc = ActiveDocument.HMIObjects.AddHMIObject("EllipseArc2", "HMIEllipseArc")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddEllipseSegment()
'VBA248
    Dim objEllipseSegment As HMIEllipseSegment
    Set objEllipseSegment = ActiveDocument.HMIObjects.AddHMIObject("EllipseSegment", "HMIEllipseSegment")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditEllipseSegment()
'VBA249
    Dim objEllipseSegment As HMIEllipseSegment
    Set objEllipseSegment = ActiveDocument.HMIObjects("EllipseSegment")
    objEllipseSegment.BorderColor = RGB(255, 0, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNameOfFirstSelectedObject()
'VBA250
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name from the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddActionToPropertyTypeCScript()
'VBA251
    Dim objEvent As HMIEvent
    Dim objCScript As HMIScriptInfo
    Dim objCircle As HMICircle
    'Create circle in the picture. If property "Radius" is changed,
    'a C-action is added:
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_AB", "HMICircle")
    Set objEvent = objCircle.Radius.Events(1)
    Set objCScript = objEvent.Actions.AddAction(hmiActionCreationTypeCScript)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowEventsOfAllObjectsInActiveDocument()
'VBA252
    Dim colEvents As HMIEvents
    Dim objEvent As HMIEvent
    Dim iMax As Integer
    Dim iIndex As Integer
    Dim iAnswer As Integer
    Dim strEventName As String
    Dim strObjectName As String
    Dim varEventType As Variant
    iIndex = 1
    iMax = ActiveDocument.HMIObjects.Count
    For iIndex = 1 To iMax
        Set colEvents = ActiveDocument.HMIObjects(iIndex).Events
        strObjectName = ActiveDocument.HMIObjects(iIndex).ObjectName
        For Each objEvent In colEvents
            strEventName = objEvent.EventName
            varEventType = objEvent.EventType
            iAnswer = MsgBox("Objectname: " & strObjectName & vbCrLf & "Eventtype: " & varEventType & vbCrLf & "Eventname: " & strEventName, vbOKCancel)
            If vbCancel = iAnswer Then Exit For
        Next objEvent
        If vbCancel = iAnswer Then Exit For
    Next iIndex
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowFolderItemsOfGlobalLibrary()
'VBA253
    Dim colFolderItems As HMIFolderItems
    Dim objFolderItem As HMIFolderItem
    Set colFolderItems = Application.SymbolLibraries(1).FolderItems
    For Each objFolderItem In colFolderItems
        MsgBox objFolderItem.Name
    Next objFolderItem
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CopyFolderItemToClipboard()
'VBA254
    Dim objGlobalLib As HMISymbolLibrary
    Set objGlobalLib = Application.SymbolLibraries(1)
    objGlobalLib.FolderItems("Folder2").Folder("Folder2").Folder.Item("Object1").CopyToClipboard
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowFolderItemsOfGlobalLibrary()
'VBA255
    Dim colFolderItems As HMIFolderItems
    Dim objFolderItem As HMIFolderItem
    Set colFolderItems = Application.SymbolLibraries(1).FolderItems
    For Each objFolderItem In colFolderItems
        MsgBox objFolderItem.Name
    Next objFolderItem
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddNewFolderToProjectLibrary()
'VBA256
    Dim objProjectLib As HMISymbolLibrary
    Set objProjectLib = Application.SymbolLibraries(2)
    objProjectLib.FolderItems.AddFolder ("My Folder")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddGraphicObject()
'VBA257
    Dim objGraphicObject As HMIGraphicObject
    Set objGraphicObject = ActiveDocument.HMIObjects.AddHMIObject("Graphic-Object", "HMIGraphicObject")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditGraphicObject()
'VBA258
    Dim objGraphicObject As HMIGraphicObject
    Set objGraphicObject = ActiveDocument.HMIObjects("Graphic-Object")
    objGraphicObject.BorderColor = RGB(255, 0, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNameOfFirstSelectedObject()
'VBA259
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name of the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DoCreateGroup()
'VBA260
    Dim objGroup As HMIGroup
    Set objGroup = ActiveDocument.Selection.CreateGroup
    objGroup.ObjectName = "Group-Object"
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditGroup()
'VBA261
    Dim objGroup As HMIGroup
    Set objGroup = ActiveDocument.HMIObjects("Group-Object")
    MsgBox objGroup.ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddGroupDisplay()
'VBA262
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("Groupdisplay", "HMIGroupDisplay")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditGroupDisplay()
'VBA263
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects("Groupdisplay")
    objGroupDisplay.BackColor = RGB(255, 0, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNameOfFirstSelectedObject()
'VBA264
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name from the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowGroupedObjectsOfFirstGroup()
'VBA265
    Dim colGroupedObjects As HMIGroupedObjects
    Dim objObject As HMIObject
    Set colGroupedObjects = ActiveDocument.HMIObjects("Group1").GroupedHMIObjects
    For Each objObject In colGroupedObjects
        MsgBox objObject.ObjectName
    Next objObject
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RemoveObjectFromGroup()
'VBA266
    Dim objGroup As HMIGroup
    Set objGroup = ActiveDocument.HMIObjects("Group1")
    objGroup.GroupedHMIObjects.Remove (1)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowDefaultObjects()
'VBA267
    Dim strType As String
    Dim strName As String
    Dim strMessage As String
    Dim iMax As Integer
    Dim iIndex As Integer
 
    iMax = Application.DefaultHMIObjects.Count
    iIndex = 1
    For iIndex = 1 To iMax
        With Application.DefaultHMIObjects(iIndex)
            strType = .Type
            strName = .ObjectName
            strMessage = strMessage & "Element: " & iIndex & " / Objecttype: " & strType & " / Objectname: " & strName
        End With
        If 0 = iIndex Mod 10 Then
            MsgBox strMessage
            strMessage = ""
        Else
            strMessage = strMessage & vbCrLf & vbCrLf
        End If
    Next iIndex
    MsgBox "Element: " & iIndex & vbCrLf & "Objecttype: " & strType & vbCrLf & "Objectname: " & strName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowFirstObjectOfCollection()
'VBA268
    Dim strName As String
    strName = ActiveDocument.HMIObjects(1).ObjectName
    MsgBox strName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DeleteObject()
'VBA269
    ActiveDocument.HMIObjects(1).Delete
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowObjectsOfDocument()
'VBA270
    Dim colObjects As HMIObjects
    Dim objObject As HMIObject
    Set colObjects = ActiveDocument.HMIObjects
    For Each objObject In colObjects
        MsgBox objObject.ObjectName
    Next objObject
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddCircle()
'VBA271
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_1", "HMICircle")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub FindObjectsByType()
'VBA272
    Dim colSearchResults As HMICollection
    Dim objMember As HMIObject
    Dim iResult As Integer
    Dim strName As String
    Set colSearchResults = ActiveDocument.HMIObjects.Find(ObjectType:="HMICircle")
    For Each objMember In colSearchResults
        iResult = colSearchResults.Count
        strName = objMember.ObjectName
        MsgBox "Found: " & CStr(iResult) & vbCrLf & "Objectname: " & strName
    Next objMember
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddIOField()
'VBA273
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IO-Field", "HMIIOField")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditIOField()
'VBA274
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects("IO-Field")
    objIOField.BorderColor = RGB(255, 0, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNameOfFirstSelectedObject()
'VBA275
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name of the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditDefaultPropertiesOfIOField()
'VBA276
    Dim objIOField As HMIIOField
    Set objIOField = Application.DefaultHMIObjects("HMIIOField")
    objIOField.BorderColor = RGB(255, 255, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowFirstObjectOfCollection()
'VBA277
    Dim strName As String
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button", "HMIButton")
    strName = objButton.LDFonts(1).Family
    MsgBox strName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowLanguageFont()
'VBA278
    Dim colLanguageFonts As HMILanguageFonts
    Dim objLanguageFont As HMILanguageFont
    Dim objButton As HMIButton
    Dim iMax As Integer
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    Set colLanguageFonts = objButton.LDFonts
    iMax = colLanguageFonts.Count
    For Each objLanguageFont In colLanguageFonts
        MsgBox "Planned fonts: " & iMax & vbCrLf & "Language-ID: " & objLanguageFont.LanguageID
    Next objLanguageFont
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ExampleForLanguageFonts()
'VBA279
    Dim colLangFonts As HMILanguageFonts
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    objButton.Text = "DefText"
    Set colLangFonts = objButton.LDFonts
    
    'Adjust fontsettings for french:
    With colLangFonts.ItemByLCID(1036)
        .Family = "Courier New"
        .Bold = True
        .Italic = False
        .Underlined = True
        .Size = 12
    End With
    'Adjust fontsettings for english:
    With colLangFonts.ItemByLCID(1033)
        .Family = "Times New Roman"
        .Bold = False
        .Italic = True
        .Underlined = False
        .Size = 14
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddLanguagesToButton()
'VBA280
    Dim objLabelText As HMILanguageText
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    '
    'Add text in actual datalanguage:
    objButton.Text = "Actual-Language Text"
    '
    'Add english text:
    Set objLabelText = ActiveDocument.HMIObjects("myButton").LDTexts.Add(1033, "English Text")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddLanguagesToButton()
'VBA281
    Dim objLabelText As HMILanguageText
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    '
    'Add text in actual datalanguage:
    objButton.Text = "Actual-Language Text"
    '
    'Add english text:
    Set objLabelText = ActiveDocument.HMIObjects("myButton").LDTexts.Add(1033, "English Text")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ConfigureSettingsOfLayer()
'VBA282
    Dim objLayer As HMILayer
    Set objLayer = ActiveDocument.Layers(1)
    With objLayer
        'configure "Layer 0"
        .MinZoom = 10
        .MaxZoom = 100
        .Name = "Configured with VBA"
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowLayer()
'VBA283
    Dim colLayers As HMILayers
    Dim objLayer As HMILayer
    Dim strLayerList As String
    Dim iCounter As Integer
    iCounter = 1
    Set colLayers = ActiveDocument.Layers
    For Each objLayer In colLayers
        If 1 = iCounter Mod 2 And 32 > iCounter Then
            strLayerList = strLayerList & vbCrLf
        ElseIf 11 > iCounter Then
            strLayerList = strLayerList & "       "
        Else
            strLayerList = strLayerList & "     "
        End If
        strLayerList = strLayerList & objLayer.Name
        iCounter = iCounter + 1
    Next objLayer
    MsgBox strLayerList
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddCircle()
'VBA284
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMISelection.SelectAll
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddLine()
'VBA285
    Dim objLine As HMILine
    Set objLine = ActiveDocument.HMIObjects.AddHMIObject("Line1", "HMILine")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditLine()
'VBA286
    Dim objLine As HMILine
    Set objLine = ActiveDocument.HMIObjects("Line1")
    objLine.BorderColor = RGB(255, 0, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNameOfFirstSelectedObject()
'VBA287
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name of the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowFirstMenuOfMenucollection()
'VBA288
    Dim strName As String
    strName = ActiveDocument.CustomMenus(1).Label
    MsgBox strName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DeleteMenu()
'VBA289
    Dim objMenu As HMIMenu
    Set objMenu = ActiveDocument.CustomMenus(1)
    objMenu.Delete
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowCustomMenusOfDocument()
'VBA290
    Dim colMenus As HMIMenus
    Dim objMenu As HMIMenu
    Dim strMenuList As String
    Set colMenus = ActiveDocument.CustomMenus
    For Each objMenu In colMenus
        strMenuList = strMenuList & objMenu.Label & vbCrLf
    Next objMenu
    MsgBox strMenuList
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub InsertApplicationSpecificMenu()
'VBA291
    Dim objMenu As HMIMenu
    Set objMenu = Application.CustomMenus.InsertMenu(1, "a_Menu1", "myApplicationMenu")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub InsertDocumentSpecificMenu()
'VBA292
    Dim objMenu As HMIMenu
    Set objMenu = ActiveDocument.CustomMenus.InsertMenu(1, "d_Menu1", "myDocumentMenu")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowFirstObjectOfCollection()
'VBA293
    Dim strName As String
    strName = ActiveDocument.CustomMenus(1).MenuItems(1).Label
    MsgBox strName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DeleteMenuItem()
'VBA294
    ActiveDocument.CustomMenus(1).MenuItems(1).Delete
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowMenuItems()
'VBA295
    Dim colMenuItems As HMIMenuItems
    Dim objMenuItem As HMIMenuItem
    Dim strItemList As String
    Set colMenuItems = ActiveDocument.CustomMenus(1).MenuItems
    For Each objMenuItem In colMenuItems
        strItemList = strItemList & objMenuItem.Label & vbCrLf
    Next objMenuItem
    MsgBox strItemList
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub InsertMenuItem()
'VBA296
    Dim objMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
    Set objMenu = ActiveDocument.CustomMenus.InsertMenu(2, "d_Menu2", "DocMenu2")
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(1, "m_Item2_1", "MenuItem 1")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowConnectorInfo_Menu()
'VBA297
    Dim objMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
    Dim strDocName As String
    strDocName = Application.ApplicationDataPath & ActiveDocument.Name
    Set objMenu = Documents(strDocName).CustomMenus.InsertMenu(1, "ConnectorMenu", "Connector_Info")
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(1, "ShowConnectInfo", "Info Connector")
End Sub
 
Sub ShowConnectorInfo()
    Dim objConnector As HMIObjConnection
    Dim iStart As Integer
    Dim iEnd As Integer
    Dim strStart As String
    Dim strEnd As String
    Dim strObjStart As String
    Dim strObjEnd As String
    Set objConnector = ActiveDocument.HMIObjects("Connector1")
    iStart = objConnector.BottomConnectedConnectionPointIndex
    iEnd = objConnector.TopConnectedConnectionPointIndex
    strObjStart = objConnector.BottomConnectedObjectName
    strObjEnd = objConnector.TopConnectedObjectName
    Select Case iStart
        Case 0
            strStart = "top"
        Case 1
            strStart = "right"
        Case 2
            strStart = "bottom"
        Case 3
            strStart = "left"
    End Select
    Select Case iEnd
        Case 0
            strEnd = "top"
        Case 1
            strEnd = "right"
        Case 2
            strEnd = "bottom"
        Case 3
            strEnd = "left"
    End Select
    MsgBox "The selected connector links the objects " & vbCrLf & "'" & strObjStart & "' and '" & strObjEnd & "'" & vbCrLf & "Connected points: " & vbCrLf & strObjStart & ": " & strStart & vbCrLf & strObjEnd & ": " & strEnd
End Sub

Private Sub Document_MenuItemClicked(ByVal MenuItem As IHMIMenuItem)
    Select Case MenuItem.Key
        Case "ShowConnectInfo"
            Call ShowConnectorInfo
    End Select
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddOLEObjectToActiveDocument()
'VBA298
    Dim objOleObject As HMIOLEObject
    Set objOleObject = ActiveDocument.HMIObjects.AddOLEObject("Wordpad Document", "Wordpad.Document.1")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditOLEObject()
'VBA299
    Dim objOleObject As HMIOLEObject
    Set objOleObject = ActiveDocument.HMIObjects("Wordpad Document")
    objOleObject.Left = 140
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNameOfFirstSelectedObject()
'VBA300
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name of the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddOptionGroup()
'VBA301
    Dim objOptionGroup As HMIOptionGroup
    Set objOptionGroup = ActiveDocument.HMIObjects.AddHMIObject("Radio-Box", "HMIOptionGroup")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditOptionGroup()
'VBA302
    Dim objOptionGroup As HMIOptionGroup
    Set objOptionGroup = ActiveDocument.HMIObjects("Radio-Box")
    objOptionGroup.BorderColor = RGB(255, 0, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNameOfFirstSelectedObject()
'VBA303
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name of the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddPictureWindow()
'VBA304
    Dim objPictureWindow As HMIPictureWindow
    Set objPictureWindow = ActiveDocument.HMIObjects.AddHMIObject("PictureWindow1", "HMIPictureWindow")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditPictureWindow()
'VBA305
    Dim objPictureWindow As HMIPictureWindow
    Set objPictureWindow = ActiveDocument.HMIObjects("PictureWindow1")
    objPictureWindow.Sizeable = True
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNameOfFirstSelectedObject()
'VBA306
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name of the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddPieSegment()
'VBA307
    Dim objPieSegment As HMIPieSegment
    Set objPieSegment = ActiveDocument.HMIObjects.AddHMIObject("PieSegment1", "HMIPieSegment")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditPieSegment()
'VBA308
    Dim objPieSegment As HMIPieSegment
    Set objPieSegment = ActiveDocument.HMIObjects("PieSegment1")
    objPieSegment.BorderColor = RGB(255, 0, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNameOfFirstSelectedObject()
'VBA309
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name of the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddPolygon()
'VBA310
    Dim objPolygon As HMIPolygon
    Set objPolygon = ActiveDocument.HMIObjects.AddHMIObject("Polygon", "HMIPolygon")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditPolygon()
    Dim objPolygon As HMIPolygon
    Set objPolygon = ActiveDocument.HMIObjects("Polygon")
    objPolygon.BorderColor = RGB (255, 0, 0)
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNameOfFirstSelectedObject()
'VBA312
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name of the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddPolyLine()
'VBA313
    Dim objPolyLine As HMIPolyLine
    Set objPolyLine = ActiveDocument.HMIObjects.AddHMIObject("PolyLine1", "HMIPolyLine")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditPolyLine()
'VBA314
    Dim objPolyLine As HMIPolyLine
    Set objPolyLine = ActiveDocument.HMIObjects("PolyLine1")
    objPolyLine.BorderColor = RGB(255, 0, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNameOfFirstSelectedObject()
'VBA315
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name of the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditDefaultPropertiesOfPolyLine()
'VBA316
    Dim objPolyLine As HMIPolyLine
    Set objPolyLine = Application.DefaultHMIObjects("HMIPolyLine")
    objPolyLine.BorderColor = RGB(255, 255, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowFirstObjectOfCollection()
'VBA317
    Dim objCircle As HMICircle
    Dim strName As String
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle", "HMICircle")
    strName = objCircle.Properties(1).Name
    MsgBox strName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DynamicToRadiusOfNewCircle()
'VBA318
    Dim objVariableTrigger As HMIVariableTrigger
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects("Circle")
    Set objVariableTrigger = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeVariableDirect, "NewDynamic1")
    objVariableTrigger.CycleType = hmiCycleType_2s
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddObject()
'VBA319
    Dim objObject As HMIObject
    Set objObject = ActiveDocument.HMIObjects.AddHMIObject("CircleAsHMIObject", "HMICircle")
'
    'Standard properties (e.g. "Position") are available every time:
    objObject.Top = 40
    objObject.Left = 40
'
    'Individual properties have to be called using
    'property "Properties":
    objObject.Properties("FlashBackColor") = True
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddRectangle()
'VBA320
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditRectangle()
'VBA321
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects("Rectangle1")
    objRectangle.BorderColor = RGB(255, 0, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNameOfFirstSelectedObject()
'VBA322
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name of the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddRoundButton()
'VBA323
    Dim objRoundButton As HMIRoundButton
    Set objRoundButton = ActiveDocument.HMIObjects.AddHMIObject("Roundbutton1", "HMIRoundButton")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditRoundButton()
'VBA324
    Dim objRoundButton As HMIRoundButton
    Set objRoundButton = ActiveDocument.HMIObjects("Roundbutton1")
    objRoundButton.BorderColor = RGB(255, 0, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNameOfFirstSelectedObject()
'VBA325
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name of the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddRoundRectangle()
'VBA326
    Dim objRoundRectangle As HMIRoundRectangle
    Set objRoundRectangle = ActiveDocument.HMIObjects.AddHMIObject("Roundrectangle1", "HMIRoundRectangle")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditRoundRectangle()
'VBA327
    Dim objRoundRectangle As HMIRoundRectangle
    Set objRoundRectangle = ActiveDocument.HMIObjects("Roundrectangle1")
    objRoundRectangle.BorderColor = RGB(255, 0, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNameOfFirstSelectedObject()
'VBA328
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name of the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicAsCSkriptToProperty()
'VBA329
    Dim objCScript As HMIScriptInfo
    Dim objCircle As HMICircle
 
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle1", "HMICircle")
    Set objCScript = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeCScript)
'
    'Define triggertype and cycletime:
    With objCScript
        .SourceCode = ""
        .Trigger.Type = hmiTriggerTypeStandardCycle
        .Trigger.CycleType = hmiCycleType_2s
        .Trigger.Name = "Trigger1"
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddActionToPropertyTypeCScript()
'VBA330
    Dim objEvent As HMIEvent
    Dim objCScript As HMIScriptInfo
    Dim objCircle As HMICircle
    'Add circle to picture. By changing of property "Radius"
    'a C-Aktion is initiated:
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_AB", "HMICircle")
    Set objEvent = objCircle.Radius.Events(1)
    Set objCScript = objEvent.Actions.AddAction(hmiActionCreationTypeCScript)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowSelectionOfDocument()
'VBA331
    Dim colSelection As HMISelectedObjects
    Dim objObject As HMIObject
    Dim strObjectList As String
    Set colSelection = ActiveDocument.Selection
    If colSelection.Count <> 0 Then
        strObjectList = "List of selected objects:"
        For Each objObject In colSelection
            strObjectList = strObjectList & vbCrLf & objObject.ObjectName
        Next objObject
    Else
        strObjectList = "No objects selected"
    End If
    MsgBox strObjectList
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SelectAllObjects()
'VBA332
    ActiveDocument.Selection.SelectAll
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddSlider()
'VBA333
    Dim objSlider As HMISlider
    Set objSlider = ActiveDocument.HMIObjects.AddHMIObject("Slider1", "HMISlider")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditSlider()
'VBA334
    Dim objSlider As HMISlider
    Set objSlider = ActiveDocument.HMIObjects("Slider1")
    objSlider.ButtonColor = RGB(255, 0, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNameOfFirstSelectedObject()
'VBA335
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name of the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DirectConnection()
'VBA336
    Dim objButton As HMIButton
    Dim objRectangleA As HMIRectangle
    Dim objRectangleB As HMIRectangle
    Dim objEvent As HMIEvent
    Dim objDirConnection As HMIDirectConnection
'
    'Add objects to active document:
    Set objRectangleA = ActiveDocument.HMIObjects.AddHMIObject("Rectangle_A", "HMIRectangle")
    Set objRectangleB = ActiveDocument.HMIObjects.AddHMIObject("Rectangle_B", "HMIRectangle")
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    With objRectangleA
        .Top = 100
        .Left = 100
    End With
    With objRectangleB
        .Top = 250
        .Left = 400
        .BackColor = RGB(255, 0, 0)
    End With
    With objButton
        .Top = 10
        .Left = 10
        .Text = "SetPosition"
    End With
'
    'Initiation of directconnection by mouseclick:
    Set objDirConnection = objButton.Events(1).Actions.AddAction(hmiActionCreationTypeDirectConnection)
    With objDirConnection
        'Sourceobjekt: Top-property of Rectangle_A
        .SourceLink.Type = hmiSourceTypeProperty
        .SourceLink.ObjectName = "Rectangle_A"
        .SourceLink.AutomationName = "Top"
'
        'Targetobjekt: Left-property of Rectangle_B
        .DestinationLink.Type = hmiDestTypeProperty
        .DestinationLink.ObjectName = "Rectangle_B"
        .DestinationLink.AutomationName = "Left"
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddStaticText()
'VBA337
    Dim objStaticText As HMIStaticText
    Set objStaticText = ActiveDocument.HMIObjects.AddHMIObject("Static_Text1", "HMIStaticText")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditStaticText()
'VBA338
    Dim objStaticText As HMIStaticText
    Set objStaticText = ActiveDocument.HMIObjects("Static_Text1")
    objStaticText.BorderColor = RGB(255, 0, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNameOfFirstSelectedObject()
'VBA339
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name of the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddStatusDisplay()
'VBA340
    Dim objStatusDisplay As HMIStatusDisplay
    Set objStatusDisplay = ActiveDocument.HMIObjects.AddHMIObject("Statusdisplay1", "HMIStatusDisplay")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditStatusDisplay()
'VBA341
    Dim objStatusDisplay As HMIStatusDisplay
    Set objStatusDisplay = ActiveDocument.HMIObjects("Statusdisplay1")
    objStatusDisplay.BorderColor = RGB(255, 0, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNameOfFirstSelectedObject()
'VBA342
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name of the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowFirstObjectOfCollection()
'VBA343
    Dim strName As String
    strName = Application.SymbolLibraries(1).Name
    MsgBox strName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowSymbolLibraries()
'VBA344
    Dim colSymbolLibraries As HMISymbolLibraries
    Dim objSymbolLibrary As HMISymbolLibrary
    Dim strLibraryList As String
    Set colSymbolLibraries = Application.SymbolLibraries
    For Each objSymbolLibrary In colSymbolLibraries
        strLibraryList = strLibraryList & objSymbolLibrary.Name & vbCrLf
    Next objSymbolLibrary
    MsgBox strLibraryList
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddTextList()
'VBA345
    Dim objTextList As HMITextList
    Set objTextList = ActiveDocument.HMIObjects.AddHMIObject("Textlist1", "HMITextList")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditTextList()
'VBA346
    Dim objTextList As HMITextList
    Set objTextList = ActiveDocument.HMIObjects("Textlist1")
    objTextList.BorderColor = RGB(255, 0, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNameOfFirstSelectedObject()
'VBA347
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name of the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowFirstObjectOfCollection()
'VBA348
    Dim strName As String
    strName = ActiveDocument.CustomToolbars(1).Key
    MsgBox strName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DeleteToolbar()
'VBA349
    Dim objToolbar As HMIToolbar
    Set objToolbar = ActiveDocument.CustomToolbars(1)
    objToolbar.Delete
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowCustomToolbarsOfDocument()
'VBA350
    Dim colToolbars As HMIToolbars
    Dim objToolbar As HMIToolbar
    Dim strToolbarList As String
    Set colToolbars = ActiveDocument.CustomToolbars
    If 0 <> colToolbars.Count Then
        For Each objToolbar In colToolbars
            strToolbarList = strToolbarList & objToolbar.Key & vbCrLf
        Next objToolbar
    Else
        strToolbarList = "No toolbars existing"
    End If
    MsgBox strToolbarList
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub InsertApplicationSpecificToolbar()
'VBA351
    Dim objToolbar As HMIToolbar
    Set objToolbar = Application.CustomToolbars.Add("a_Toolbar1")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub InsertDocumentSpecificToolbar()
'VBA352
    Dim objToolbar As HMIToolbar
    Set objToolbar = ActiveDocument.CustomToolbars.Add("d_Toolbar1")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowFirstObjectOfCollection()
'VBA353
    Dim strType As String
    strType = ActiveDocument.CustomToolbars(1).ToolbarItems(1).ToolbarItemType
    MsgBox strType
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DeleteToolbarItem()
'VBA354
    ActiveDocument.CustomToolbars(1).ToolbarItems(1).Delete
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowToolbarItems()
'VBA355
    Dim colToolbarItems As HMIToolbarItems
    Dim objToolbarItem As HMIToolbarItem
    Dim strTypeList As String
    Set colToolbarItems = ActiveDocument.CustomToolbars(1).ToolbarItems
    If 0 <> colToolbarItems.Count Then
        For Each objToolbarItem In colToolbarItems
            strTypeList = strTypeList & objToolbarItem.ToolbarItemType & vbCrLf
        Next objToolbarItem
    Else
        strTypeList = "No Toolbaritems existing"
    End If
    MsgBox strTypeList
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub InsertToolbarItem()
'VBA356
    Dim objToolbar As HMIToolbar
    Dim objToolbarItem As HMIToolbarItem
    Set objToolbar = ActiveDocument.CustomToolbars.Add("d_Toolbar2")
    Set objToolbarItem = objToolbar.ToolbarItems.InsertToolbarItem(1, "t_Item2_1", "ToolbarItem 1")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicAsVBSkriptToProperty()
'VBA357
    Dim objVBScript As HMIScriptInfo
    Dim objCircle As HMICircle
     
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle1", "HMICircle")
    Set objVBScript = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeVBScript)
'
    'Define cycletime and sourcecode
    With objVBScript
        .SourceCode = ""
        .Trigger.Type = hmiTriggerTypeStandardCycle
        .Trigger.CycleType = hmiCycleType_2s
        .Trigger.Name = "Trigger1"
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeAnalog()
'VBA358
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeAnalog
        .AnalogResultInfos.ElseCase = 200
        'Activate variable-statecheck
        .VariableStateChecked = True
    End With
    With objDynDialog.VariableStateValues(1)
        'define a value for every state:
        .VALUE_ACCESS_FAULT = 20
        .VALUE_ADDRESS_ERROR = 30
        .VALUE_CONVERSION_ERROR = 40
        .VALUE_HANDSHAKE_ERROR = 60
        .VALUE_HARDWARE_ERROR = 70
        .VALUE_INVALID_KEY = 80
        .VALUE_MAX_LIMIT = 90
        .VALUE_MAX_RANGE = 100
        .VALUE_MIN_LIMIT = 110
        .VALUE_MIN_RANGE = 120
        .VALUE_NOT_ESTABLISHED = 130
        .VALUE_SERVERDOWN = 140
        .VALUE_STARTUP_VALUE = 150
        .VALUE_TIMEOUT = 160
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicAsVariableDirectToProperty()
'VBA359
    Dim objVariableTrigger As HMIVariableTrigger
    Dim objCircle As HMICircle
 
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle1", "HMICircle")
    Set objVariableTrigger = objCircle.Top.CreateDynamic(hmiDynamicCreationTypeVariableDirect, "NewDynamic1")
'
    'Define cycletime
    With objVariableTrigger
        .CycleType = hmiCycleType_2s
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DynamicWithVariableTriggerCycle()
'VBA360
    Dim objVBScript As HMIScriptInfo
    Dim objVarTrigger As HMIVariableTrigger
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_VariableTrigger", "HMICircle")
    Set objVBScript = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeVBScript)
    With objVBScript
        'Definition of triggername and cycletime is to do with the Add-methode
        Set objVarTrigger = .Trigger.VariableTriggers.Add("VarTrigger", hmiVariableCycleType_10s)
        .SourceCode = ""
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNumberOfExistingViews()
'VBA361
    Dim iMaxViews As Integer
    iMaxViews = ActiveDocument.Views.Count
    MsgBox "Number of copies from active document: " & iMaxViews
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddView()
'VBA362
    Dim objView As HMIView
    Set objView = ActiveDocument.Views.Add
    objView.Activate
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowNumberOfExistingViews()
'VBA363
    Dim iMaxViews As Integer
    iMaxViews = ActiveDocument.Views.Count
    MsgBox "Number of copies from active document: " & iMaxViews
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddViewToActiveDocument()
'VBA364
    Dim objView As HMIView
    Set objView = ActiveDocument.Views.Add
    objView.Activate
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateVBActionToClickedEvent()
'VBA365
    Dim objButton As HMIButton
    Dim objCircle As HMICircle
    Dim objEvent As HMIEvent
    Dim objVBScript As HMIScriptInfo
    Dim strCode As String
    strCode = "Dim myCircle" & vbCrLf & "Set myCircle = "
    strCode = strCode & "HMIRuntime.ActiveScreen.ScreenItems(""Circle_VB"")"
    strCode = strCode & vbCrLf & "myCircle.Radius = myCircle.Radius + 5"
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_VB", "HMICircle")
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    With objCircle
        .Top = 100
        .Left = 100
        .BackColor = RGB(255, 0, 0)
    End With
    With objButton
        .Top = 10
        .Left = 10
        .Width = 120
        .Text = "Increase Radius"
    End With
    'Define event and assign sourcecode:
    Set objVBScript = objButton.Events(1).Actions.AddAction(hmiActionCreationTypeVBScript)
    With objVBScript
        .SourceCode = strCode
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateMenuItem()
'VBA366
    Dim objMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
'
    'Create new menu "Delete Objects":
    Set objMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DeleteObjects", "Delete Objects")
'
    'Add two menuitems to the menu "Delete Objects
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(1, "DeleteAllRectangles", "Delete Rectangles")
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(2, "DeleteAllCircles", "Delete Circles")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ActiveDocumentConfiguration()
'VBA367
    Application.ActiveDocument.Views.Add
    Application.ActiveDocument.Views(1).ActiveLayer = 2
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub PolygonCoordinatesOutput()
'VBA369
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



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub LineAdd()
'VBA370
    Dim objLine As HMILine
    Set objLine = ActiveDocument.HMIObjects.AddHMIObject("myLine", "HMILine")
    With objLine
        .BorderColor = RGB(255, 0, 0)
        .index = hmiLineIndexTypeStartPoint
        .ActualPointLeft = 12
        .ActualPointTop = 34
        .index = hmiLineIndexTypeEndPoint
        .ActualPointLeft = 74
        .ActualPointTop = 64
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub LineAdd()
'VBA371
    Dim objLine As HMILine
    Set objLine = ActiveDocument.HMIObjects.AddHMIObject("myLine", "HMILine")
    With objLine
        .BorderColor = RGB(255, 0, 0)
        .index = hmiLineIndexTypeStartPoint
        .ActualPointLeft = 12
        .ActualPointTop = 34
        .index = hmiLineIndexTypeEndPoint
        .ActualPointLeft = 74
        .ActualPointTop = 64
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub IOFieldConfiguration()
'VBA372
    Dim objIOField As HMIIOField
'
    'Add new IO-Feld to active document:
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .AdaptBorder = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub PictureWindowConfig()
'VBA373
    Dim objPicWindow As HMIPictureWindow
'
    'Add new picturewindow to active document:
    Set objPicWindow = ActiveDocument.HMIObjects.AddHMIObject("PicWindow1", "HMIPictureWindow")
    With objPicWindow
        .AdaptPicture = False
        .AdaptSize = False
        .Caption = True
        .CaptionText = "Picturewindow in runtime"
        .OffsetLeft = 5
        .OffsetTop = 10
'
        'Replace the picturename "Test.PDL" with the name of
        'an existing document from your "GraCS"-Folder of your active project
        .PictureName = "Test.PDL"
        .ScrollBars = True
        .ServerPrefix = ""
        .TagPrefix = "Struct."
        .UpdateCycle = 5
        .Zoom = 100
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub PictureWindowConfig()
'VBA374
    Dim objPicWindow As HMIPictureWindow
'
    'Add new picturewindow to active document:
    Set objPicWindow = ActiveDocument.HMIObjects.AddHMIObject("PicWindow1", "HMIPictureWindow")
    With objPicWindow
        .AdaptPicture = False
        .AdaptSize = False
        .Caption = True
        .CaptionText = "Picturewindow in runtime"
        .OffsetLeft = 5
        .OffsetTop = 10
'
        'Replace the picturename "Test.PDL" with the name of
        'an existing document from your "GraCS"-Folder of your active project
        .PictureName = "Test.PDL"
        .PictureName = "Testpicture.BMP"
        .ScrollBars = True
        .ServerPrefix = ""
        .TagPrefix = "Struct."
        .UpdateCycle = 5
        .Zoom = 100
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA375
    Dim objBarGraph As HMIBarGraph
'
    'Add new BarGraph to active document:
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolut
        .TypeAlarmHigh = False
        'Activate monitoring
        .CheckAlarmHigh = True
        'Set barcolor to "yellow"
        .ColorAlarmHigh = RGB(255, 255, 0)
        'set upper limit to "50"
        .AlarmHigh = 50
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA376
    Dim objBarGraph As HMIBarGraph
'
    'Add new bargraph to active document:
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolut
        .TypeAlarmLow = False
        'Activate monitoring
        .CheckAlarmLow = True
        'Set Barcolor to "yellow"
        .ColorAlarmLow = RGB(255, 255, 0)
        'set lower limit to "10"
        .AlarmLow = 10
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphConfiguration()
'VBA377
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .Alignment = True
        .Scaling = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub IOFieldConfiguration()
'VBA378
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .AlignmentLeft = 1
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub IOFieldConfiguration()
'VBA379
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .AlignmentLeft = 1
        .AlignmentTop = 1
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA380
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        'Depth-angle a = 15 degrees
        .AngleAlpha = 15
        'Depth-angle b = 45 degrees
        .AngleBeta = 45
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA381
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        'Depth-angle a = 15 degrees
        .AngleAlpha = 15
        'Depth-angle b = 45 degrees
        .AngleBeta = 45
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateExcelApplication()
'VBA382
'
    'Open Excel invisible
    Dim objExcelApp As New Excel.Application
    MsgBox objExcelApp
    'Delete the reference to Excel and close it
    Set objExcelApp = Nothing
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowApplicationDataPath()
'VBA383
    MsgBox Application.ApplicationDataPath
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub IOFieldConfiguration()
'VBA385
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .AssumeOnExit = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub IOFieldConfiguration()
'VBA386
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .AssumeOnFull = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DirectConnection()
'VBA387
    Dim objButton As HMIButton
    Dim objRectangleA As HMIRectangle
    Dim objRectangleB As HMIRectangle
    Dim objEvent As HMIEvent
    Dim objDynConnection As HMIDirectConnection
'
    'Add objects to active document:
    Set objRectangleA = ActiveDocument.HMIObjects.AddHMIObject("Rectangle_A", "HMIRectangle")
    Set objRectangleB = ActiveDocument.HMIObjects.AddHMIObject("Rectangle_B", "HMIRectangle")
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
'
    'to position and configure objects:
    With objRectangleA
        .Top = 100
        .Left = 100
    End With
    With objRectangleB
        .Top = 250
        .Left = 400
        .BackColor = RGB(255, 0, 0)
    End With
    With objButton
        .Top = 10
        .Left = 10
        .Text = "SetPosition"
    End With
'
    'Directconnection is initiate by mouseclick:
    Set objDynConnection = objButton.Events(1).Actions.AddAction(hmiActionCreationTypeDirectConnection)
    With objDynConnection
        'Sourceobject: Top-Property of Rectangle_A
        .SourceLink.Type = hmiSourceTypeProperty
        .SourceLink.ObjectName = "Rectangle_A"
        .SourceLink.AutomationName = "Top"
'
        'Targetobject: Left-Property of Rectangle_B
        .DestinationLink.Type = hmiDestTypeProperty
        .DestinationLink.ObjectName = "Rectangle_B"
        .DestinationLink.AutomationName = "Left"
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub OutputDataLanguages()
'VBA388
    Dim colDataLang As HMIDataLanguages
    Dim objDataLang As HMIDataLanguage
    Dim strLangList As String
    Dim iCounter As Integer
'
    'Save collection of datalanguages
    'into variable "colDataLang"
    Set colDataLang = Application.AvailableDataLanguages
    iCounter = 1
'
    'Get every languagename and the assigned ID
    For Each objDataLang In colDataLang
        With objDataLang
            If 0 = iCounter Mod 3 Or 1 = iCounter Then
                strLangList = strLangList & vbCrLf & .LanguageID & " " & .LanguageName
            Else
                strLangList = strLangList & " / " & .LanguageID & " " & .LanguageName
            End If
        End With
        iCounter = iCounter + 1
    Next objDataLang
    MsgBox strLangList
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphConfiguration()
'VBA389
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .Average = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA390
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Axe = 1
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphConfiguration()
'VBA391
    Dim objBar As HMIBarGraph
    Set objBar = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBar
        .AxisSection = 1
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ButtonConfiguration()
'VBA392
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .BackBorderWidth = 2
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RectangleConfiguration()
'VBA393
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .BackColor = RGB(255, 255, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphConfiguration()
'VBA394
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .BackColor2 = RGB(255, 255, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphConfiguration()
'VBA395
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .BackColor3 = RGB(0, 0, 255)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SliderConfiguration()
'VBA396
    Dim objSlider As HMISlider
    Set objSlider = ActiveDocument.HMIObjects.AddHMIObject("SliderObject1", "HMISlider")
    With objSlider
        .BackColorBottom = RGB(0, 0, 255)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SliderConfiguration()
'VBA397
    Dim objSlider As HMISlider
    Set objSlider = ActiveDocument.HMIObjects.AddHMIObject("SliderObject1", "HMISlider")
    With objSlider
        .BackColorTop = RGB(255, 255, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RectangleConfiguration()
'VBA398
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .BackFlashColorOff = RGB(255, 255, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RectangleConfiguration()
'VBA399
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .BackFlashColorOn = RGB(0, 0, 255)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA400
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Background = False
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA401
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .BarDepth = 40
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA402
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .BarHeight = 60
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA403
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .BarWidth = 80
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub StatusDisplayConfiguration()
'VBA404
    Dim objStatDisp As HMIStatusDisplay
    Set objStatDisp = ActiveDocument.HMIObjects.AddHMIObject("Statusdisplay1", "HMIStatusDisplay")
    With objStatDisp
        .BasePicReferenced = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub StatusDisplayConfiguration()
'VBA405
    Dim objStatDisp As HMIStatusDisplay
    Set objStatDisp = ActiveDocument.HMIObjects.AddHMIObject("Statusdisplay1", "HMIStatusDisplay")
    With objStatDisp
        .BasePicTransColor = RGB(255, 255, 0)
        .BasePicUseTransColor = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub StatusDisplayConfiguration()
'VBA406
    Dim objStatDisp As HMIStatusDisplay
    Set objStatDisp = ActiveDocument.HMIObjects.AddHMIObject("Statusdisplay1", "HMIStatusDisplay")
    With objStatDisp
'
        'To use this example copy a Bitmap-Graphic
        'to the "GraCS"-Folder of the actual project.
        'Replace the picturename "Testpicture.BMP" with the name of
        'the picture you copied
        .BasePicture = "Testpicture.BMP"
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub StatusDisplayConfiguration()
'VBA407
    Dim objStatDisp As HMIStatusDisplay
    Set objStatDisp = ActiveDocument.HMIObjects.AddHMIObject("Statusdisplay1", "HMIStatusDisplay")
    With objStatDisp
        .BasePicTransColor = RGB(255, 255, 0)
        .BasePicUseTransColor = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA408
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .BaseX = 80
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA409
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .BaseY = 100
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeBit()
'VBA410
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_B", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeBit
        .Trigger.VariableTriggers(1).CycleType = hmiVariableCycleType_5s
        .BitResultInfo.BitNumber = 1
        .BitResultInfo.BitSetValue = 40
        .BitResultInfo.BitNotSetValue = 80
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeBit()
'VBA411
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_B", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeBit
        .BitResultInfo.BitNumber = 1
        .BitResultInfo.BitSetValue = 40
        .BitResultInfo.BitNotSetValue = 80
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeBit()
'VBA412
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_B", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeBit
        .BitResultInfo.BitNumber = 1
        .BitResultInfo.BitSetValue = 40
        .BitResultInfo.BitNotSetValue = 80
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ExampleForLanguageFonts()
'VBA413
    Dim colLangFonts As HMILanguageFonts
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    objButton.Text = "Displaytext"
    Set colLangFonts = objButton.LDFonts
    'Set french fontproperties:
    With colLangFonts.ItemByLCID(1036)
        .Family = "Courier New"
        .Bold = True
        .Italic = False
        .Underlined = True
        .Size = 12
    End With
    'Set english fontproperties:
    With colLangFonts.ItemByLCID(1033)
        .Family = "Times New Roman"
        .Bold = False
        .Italic = True
        .Underlined = False
        .Size = 14
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ApplicationWindowConfig()
'VBA414
    Dim objAppWindow As HMIApplicationWindow
'
    'Add new applicationwindow to active document:
    Set objAppWindow = ActiveDocument.HMIObjects.AddHMIObject("AppWindow1", "HMIApplicationWindow")
    With objAppWindow
        .WindowBorder = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RectangleConfiguration()
'VBA415
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .BorderBackColor = RGB(255, 255, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RectangleConfiguration()
'VBA416
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .BorderColor = RGB(0, 0, 255)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ButtonConfiguration()
'VBA417
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .BorderColorBottom = RGB(255, 0, 0)
        .BorderColorTop = RGB(0, 0, 255)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ButtonConfiguration()
'VBA418
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .BorderColorBottom = RGB(255, 0, 0)
        .BorderColorTop = RGB(0, 0, 255)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub LineConfiguration()
'VBA419
    Dim objLine As HMILine
    Set objLine = ActiveDocument.HMIObjects.AddHMIObject("Line1", "HMILine")
    With objLine
        .BorderEndStyle = 393219
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RectangleConfiguration()
'VBA420
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .BorderFlashColorOff = RGB(0, 0, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RectangleConfiguration()
'VBA421
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .BorderFlashColorOn = RGB(255, 0, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RectangleConfiguration()
'VBA422
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .BorderStyle = 1
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CircleConfiguration()
'VBA423
    Dim objCircle As IHMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle1", "HMICircle")
    With objCircle
        .BorderWidth = 2
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateOptionGroup()
'VBA424
    Dim objRadioBox As HMIOptionGroup
    Dim iCounter As Integer
    Set objRadioBox = ActiveDocument.HMIObjects.AddHMIObject("RadioBox_1", "HMIOptionGroup")
    iCounter = 1
    With objRadioBox
        .Height = 100
        .Width = 180
        .BoxCount = 4
        .BoxAlignment = False
        For iCounter = 1 To .BoxCount
            .index = iCounter
            .Text = "CustomText" & .index
        Next iCounter
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateOptionGroup()
'VBA425
    Dim objRadioBox As HMIOptionGroup
    Dim iCounter As Integer
    Set objRadioBox = ActiveDocument.HMIObjects.AddHMIObject("RadioBox_1", "HMIOptionGroup")
    iCounter = 1
    With objRadioBox
        .Height = 100
        .Width = 180
        .BoxCount = 4
        .BoxAlignment = True
        For iCounter = 1 To .BoxCount
            .index = iCounter
            .Text = "CustomText" & .index
        Next iCounter
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub IOFieldConfiguration()
'VBA426
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .BoxType = 1
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA427
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .Button1Width = 50
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA428
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .Button2Width = 50
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA429
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .Button3Width = 50
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA430
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .Button4Width = 50
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SliderConfiguration()
'VBA431
    Dim objSlider As HMISlider
    Set objSlider = ActiveDocument.HMIObjects.AddHMIObject("SliderObject1", "HMISlider")
    With objSlider
        .ButtonColor = RGB(255, 255, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ApplicationWindowConfig()
'VBA432
    Dim objAppWindow As HMIApplicationWindow
    Set objAppWindow = ActiveDocument.HMIObjects.AddHMIObject("AppWindow", "HMIApplicationWindow")
    With objAppWindow
        .Caption = True
        .CloseButton = False
        .Height = 200
        .Left = 10
        .MaximizeButton = True
        .Moveable = False
        .OnTop = True
        .Sizeable = True
        .Top = 20
        .Visible = True
        .Width = 250
        .WindowBorder = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub PictureWindowConfig()
'VBA433
    Dim objPicWindow As HMIPictureWindow
    Set objPicWindow = ActiveDocument.HMIObjects.AddHMIObject("PicWindow1", "HMIPictureWindow")
    With objPicWindow
        .AdaptPicture = False
        .AdaptSize = False
        .Caption = True
        .CaptionText = "Picturewindow in runtime"
        .OffsetLeft = 5
        .OffsetTop = 10
        'Replace the picturename "Test.PDL" with the name of
        'an existing document from your "GraCS"-Folder of your active project
        .PictureName = "Test.PDL"
        .ScrollBars = True
        .ServerPrefix = ""
        .TagPrefix = "Struct."
        .UpdateCycle = 5
        .Zoom = 100
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA434
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeAlarmHigh = False
        'Activate monitoring
        .CheckAlarmHigh = True
        'Set barcolor to "yellow"
        .ColorAlarmHigh = RGB(255, 255, 0)
        'Set upper limit to "50"
        .AlarmHigh = 50
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA435
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeAlarmLow = False
        'Activate monitoring
        .CheckAlarmLow = True
        'Set barcolor to "yellow"
        .ColorAlarmLow = RGB (255, 255, 0)
        'Set lower limit to "10"
        .AlarmLow = 10
    End With
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateMenuItem()
'VBA436
    Dim objMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
'
    'Add new menu "Delete objects" to menubar:
    Set objMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DeleteObjects", "Delete objects")
'
    'Add two menuitems to the new menu
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(1, "DeleteAllRectangles", "Delete Rectangles")
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(2, "DeleteAllCircles", "Delete Circles")
    With objMenu.MenuItems
        .Item("DeleteAllRectangles").Checked = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA437
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeLimitHigh4 = False
        'Activate monitoring
        .CheckLimitHigh4 = True
        'set barcolor to "red"
        .ColorLimitHigh4 = RGB(255, 0, 0)
        'Set upper limit to "70"
        .LimitHigh4 = 70
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA438
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeLimitHigh5 = False
        'Activate monitoring
        .CheckLimitHigh5 = True
        'set barcolor to "black"
        .ColorLimitHigh5 = RGB(0, 0, 0)
        'Set upper limit to "80"
        .LimitHigh5 = 80
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA439
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeLimitLow4 = False
        'Activate monitoring
        .CheckLimitLow4 = True
        'Set barcolor to "green"
        .ColorLimitLow4 = RGB(0, 255, 0)
        'set lower limit to "5"
        .LimitLow4 = 5
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA440
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeLimitLow5 = False
        'Activate monitoring
        .CheckLimitLow5 = True
        'Set barcolor to "white"
        .ColorLimitLow5 = RGB(255, 255, 255)
        'set lower limit to "0"
        .LimitLow5 = 0
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA441
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeToleranceHigh = False
        'Activate monitoring
        .CheckToleranceHigh = True
        'Set barcolor to "yellow"
        .ColorToleranceHigh = RGB(255, 255, 0)
        'Set upper limit to "45"
        .ToleranceHigh = 45
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA442
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeToleranceLow = False
        'Activate monitoring
        .CheckToleranceLow = True
        'Set barcolor to "yellow"
        .ColorToleranceLow = RGB(255, 255, 0)
        'Set lower limit to "15"
        .ToleranceLow = 15
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA443
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeWarningHigh = False
        'Activate monitoring
        .CheckWarningHigh = True
        'Set barcolor to "red"
        .ColorWarningHigh = RGB(255, 0, 0)
        'Set upper limit to "75"
        .WarningHigh = 75
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA444
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeWarningLow = False
        'Activate monitoring
        .CheckWarningLow = True
        'Set barcolor to "magenta"
        .ColorWarningLow = RGB(255, 0, 255)
        'Set lower limit to "12"
        .WarningLow = 12
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub IOFieldConfiguration()
'VBA445
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .ClearOnError = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub IOFieldConfiguration()
'VBA446
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .ClearOnNew = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ApplicationWindowConfig()
'VBA447
    Dim objAppWindow As HMIApplicationWindow
    Set objAppWindow = ActiveDocument.HMIObjects.AddHMIObject("AppWindow1", "HMIApplicationWindow")
    With objAppWindow
        .CloseButton = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA449
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeAlarmHigh = False
        'Activate monitoring
        .CheckAlarmHigh = True
        'Set barcolor to "red"
        .ColorAlarmHigh = RGB(255, 0, 0)
        'Set upper limit to "50"
        .AlarmHigh = 50
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA450
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeAlarmLow = False
        'Activate monitoring
        .CheckAlarmLow = True
        'Set barcolor to "red"
        .ColorAlarmLow = RGB(255, 0, 0)
        'Set lower limit to "10"
        .AlarmLow = 10
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SliderConfiguration()
'VBA451
    Dim objSlider As HMISlider
    Set objSlider = ActiveDocument.HMIObjects.AddHMIObject("SliderObject1", "HMISlider")
    With objSlider
        .ColorBottom = RGB(255, 0, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA452
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .ColorChangeType = False
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA453
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeLimitHigh4 = False
        'Activate monitoring
        .CheckLimitHigh4 = True
        'Set barcolor to "red"
        .ColorLimitHigh4 = RGB(255, 0, 0)
        'Set upper limit to "70"
        .LimitHigh4 = 70
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA454
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeLimitHigh5 = False
        'Activate monitoring
        .CheckLimitHigh5 = True
        'Set barcolor to "black"
        .ColorLimitHigh5 = RGB(0, 0, 0)
        'Set upper limit to "80"
        .LimitHigh5 = 80
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA455
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeLimitLow4 = False
        'Activate monitoring
        .CheckLimitLow4 = True
        'Set barcolor to "green"
        .ColorLimitLow4 = RGB(0, 255, 0)
        'Set lower limit to "5"
        .LimitLow4 = 5
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA456
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeLimitLow5 = False
        'Activate monitoring
        .CheckLimitLow5 = True
        'Set barcolor to "white"
        .ColorLimitLow5 = RGB(255, 255, 255)
        'Set lower limit to "0"
        .LimitLow5 = 0
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA457
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeToleranceHigh = False
        'Activate monitoring
        .CheckToleranceHigh = True
        'Set barcolor to "yellow"
        .ColorToleranceHigh = RGB(255, 255, 0)
        'Set upper limit to "45"
        .ToleranceHigh = 45
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA458
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeToleranceLow = False
        'Activate monitoring
        .CheckToleranceLow = True
        'Set barcolor to "yellow"
        .ColorToleranceLow = RGB(255, 255, 0)
        'Set lower limit to "15"
        .ToleranceLow = 15
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SliderConfiguration()
'VBA459
    Dim objSlider As HMISlider
    Set objSlider = ActiveDocument.HMIObjects.AddHMIObject("SliderObject1", "HMISlider")
    With objSlider
        .ColorTop = RGB(255, 128, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA460
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeWarningHigh = False
        'Activate monitoring
        .CheckWarningHigh = True
        'Set barcolor to "red"
        .ColorWarningHigh = RGB(255, 0, 0)
        'Set upper limit to "75"
        .WarningHigh = 75
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA461
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeWarningLow = False
        'Activate monitoring
        .CheckWarningLow = True
        'Set barcolor to "magenta"
        .ColorWarningLow = RGB(255, 0, 255)
        'Set lower limit to "12"
        .WarningLow = 12
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub Document_Opened(CancelForwarding As Boolean)
'VBA462
    MsgBox Application.Commandline
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub IncreaseCircleRadiusWithVBScript()
'VBA463
    Dim objButton As HMIButton
    Dim objCircleA As HMICircle
    Dim objEvent As HMIEvent
    Dim objVBScript As HMIScriptInfo
    Dim strCode As String
    strCode = "Dim objCircle" & vbCrLf & "Set objCircle = "
    strCode = strCode & "hmiRuntime.ActiveScreen.ScreenItems(""CircleVB"")"
    strCode = strCode & vbCrLf & "objCircle.Radius = objCircle.Radius + 5"
    Set objCircleA = ActiveDocument.HMIObjects.AddHMIObject("CircleVB", "HMICircle")
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    With objCircleA
        .Top = 100
        .Left = 100
    End With
    With objButton
        .Top = 10
        .Left = 10
        .Width = 200
        .Text = "Increase Radius"
    End With
    'On every mouseclick the radius will be increased:
    Set objEvent = objButton.Events(1)
    Set objVBScript = objButton.Events(1).Actions.AddAction(hmiActionCreationTypeVBScript)
    objVBScript.SourceCode = strCode
    Select Case objVBScript.Compiled
        Case True
            MsgBox "Compilation OK!"
        Case False
            MsgBox "Errors by compilation!"
    End Select
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowConfigurationFileName()
'VBA464
    MsgBox ActiveDocument.Application.ConfigurationFileName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ObjectsInActiveDocument()
'VBA465
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim objDocument As Document
    Set objDocument = Application.Documents.Add hmiDocumentTypeVisible
    Dim iIndex As Integer
    iIndex = 1
    For iIndex = 1 To 5
        Set objCircle = objDocument.HMIObjects.AddHMIObject("Circle" & iIndex, "HMICircle")
        Set objRectangle = objDocument.HMIObjects.AddHMIObject("Rectangle" & iIndex, "HMIRectangle")
        With objCircle
            .Top = (10 * iIndex)
            .Left = (10 * iIndex)
        End With
        With objRectangle
            .Top = ((10 * iIndex) + 50)
            .Left = (10 * iIndex)
        End With
    Next iIndex
    MsgBox "There are " & objDocument.HMIObjects.Count & " objects in the document"
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowDataLanguage()
'VBA466
    MsgBox Application.CurrentDataLanguage
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowDesktopLanguage()
'VBA467
    MsgBox Application.CurrentDesktopLanguage
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub IOFieldConfiguration()
'VBA468
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    Application.ActiveDocument.CursorMode = True
    With objIOField
        .CursorControl = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ActiveDocumentConfiguration()
'VBA469
    Application.ActiveDocument.CursorMode = True
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowCustomMenuInformation()
'VBA470
    Dim strKey As String
    Dim strLabel As String
    Dim strOutput As String
    Dim iIndex As Integer
    For iIndex = 1 To ActiveDocument.CustomMenus.Count
        strKey = ActiveDocument.CustomMenus(iIndex).Key
        strLabel = ActiveDocument.CustomMenus(iIndex).Label
        strOutput = strOutput & vbCrLf & "Key: " & strKey & "  Label: " & strLabel
    Next iIndex
    If 0 = ActiveDocument.CustomMenus.Count Then
        strOutput = "There are no custommenus for the document created."
    End If
    MsgBox strOutput
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowCustomToolbarInformation()
'VBA471
    Dim strKey As String
    Dim strOutput As String
    Dim iIndex As Integer
    For iIndex = 1 To ActiveDocument.CustomToolbars.Count
        strKey = ActiveDocument.CustomToolbars(iIndex).Key
        strOutput = strOutput & vbCrLf & "Key: " & strKey
    Next iIndex
    If 0 = ActiveDocument.CustomToolbars.Count Then
        strOutput = "There are no toolbars created for this document."
    End If
    MsgBox strOutput
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DynamicToRadiusOfNewCircle()
'VBA474
    Dim objCircle As hmiCircle
    Dim VariableTrigger As HMIVariableTrigger
    Set objCircle = Application.ActiveDocument.HMIObjects.AddHMIObject("Circle1", "HMICircle")
    Set VariableTrigger = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeVariableDirect, "NewDynamic1")
    VariableTrigger.CycleType = hmiVariableCycleType_2s
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub IOFieldConfiguration()
'VBA475
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .DataFormat = 1
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowDefaultObjectNames()
'VBA476
    Dim strOutput As String
    Dim iIndex As Integer
    For iIndex = 1 To Application.DefaultHMIObjects.Count
        strOutput = strOutput & vbCrLf & Application.DefaultHMIObjects(iIndex).ObjectName
    Next iIndex
    MsgBox strOutput
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DirectConnection()
'VBA477
    Dim objButton As HMIButton
    Dim objRectangleA As HMIRectangle
    Dim objRectangleB As HMIRectangle
    Dim objEvent As HMIEvent
    Dim objDirConnection As HMIDirectConnection
    Set objRectangleA = ActiveDocument.HMIObjects.AddHMIObject("Rectangle_A", "HMIRectangle")
    Set objRectangleB = ActiveDocument.HMIObjects.AddHMIObject("Rectangle_B", "HMIRectangle")
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    With objRectangleA
        .Top = 100
        .Left = 100
    End With
    With objRectangleB
        .Top = 250
        .Left = 400
        .BackColor = RGB(255, 0, 0)
    End With
    With objButton
        .Top = 10
        .Left = 10
        .Width = 100
        .Text = "SetPosition"
    End With
'
    'Directconnection is initiated by mouseclick:
    Set objDirConnection = objButton.Events(1).Actions.AddAction(hmiActionCreationTypeDirectConnection)
    With objDirConnection
        'Sourceobject: Property "Top" of Rectangle_A
        .SourceLink.Type = hmiSourceTypeProperty
        .SourceLink.ObjectName = "Rectangle_A"
        .SourceLink.AutomationName = "Top"
'
        'Targetobject: Property "Left" of Rectangle_B
        .DestinationLink.Type = hmiDestTypeProperty
        .DestinationLink.ObjectName = "Rectangle_B"
        .DestinationLink.AutomationName = "Left"
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SliderConfiguration()
'VBA478
    Dim objSlider As HMISlider
    Set objSlider = ActiveDocument.HMIObjects.AddHMIObject("SliderObject1", "HMISlider")
    With objSlider
        .Direction = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DisableVBAEvents()
'VBA479
    Application.DisableVBAEvents = False
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowAllObjectDisplayNames()
'VBA480
    Dim strOutput As String
    Dim iIndex1 As Integer
    iIndex1 = 1
    strOutput = "List of all properties-displaynames from object """ & Application.DefaultHMIObjects(1).ObjectName & """" & vbCrLf & vbCrLf
    For iIndex1 = 1 To Application.DefaultHMIObjects(1).Properties.Count
        strOutput = strOutput & Application.DefaultHMIObjects(1).Properties(iIndex1).DisplayName & " / "
    Next iIndex1
    MsgBox strOutput
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowLabelTexts()
'VBA481
    Dim objLangText As HMILanguageText
    Dim iIndex As Integer
    For iIndex = 1 To ActiveDocument.CustomMenus(1).LDLabelTexts.Count
        Set objLangText = ActiveDocument.CustomMenus(1).LDLabelTexts(iIndex)
        MsgBox objLangText.DisplayName
    Next iIndex
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowDocuments()
'VBA482
    Dim colDocuments As Documents
    Dim objDocument As Document
    Dim strOutput As String
    Set colDocuments = Application.Documents
    strOutput = "List of all opened documents:" & vbCrLf
    For Each objDocument In colDocuments
        strOutput = strOutput & vbCrLf & objDocument.Name
    Next objDocument
    MsgBox strOutput
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowPropertiesDynamicsofAllObjects()
'VBA483
    Dim objObject As HMIObject
    Dim colObjects As HMIObjects
    Dim colProperties As HMIProperties
    Dim objProperty As HMIProperty
    Dim strOutput As String
    Set colObjects = Application.ActiveDocument.HMIObjects
    For Each objObject In colObjects
        Set colProperties = objObject.Properties
        For Each objProperty In colProperties
            If 0 <> objProperty.DynamicStateType Then
                strOutput = strOutput & vbCrLf & objObject.ObjectName & " - " & objProperty.DisplayName & ": Statetype " & objProperty.Dynamic.DynamicStateType
            End If
        Next objProperty
    Next objObject
    MsgBox strOutput
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub IOFieldConfiguration()
'VBA484
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .EditAtOnce = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeAnalog()
'VBA485
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeAnalog
        .AnalogResultInfos.Add 50, 40
        .AnalogResultInfos.Add 100, 80
        .AnalogResultInfos.ElseCase = 100
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DisableMenuItem()
'VBA486
    Dim objMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
'
    'Add a new menu "Delete objects"
    Set objMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DeleteObjects", "Delete objects")
'
    'Add two menuitems to the new menu
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(1, "DeleteAllRectangles", "Delete rectangles")
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(2, "DeleteAllCircles", "Delete circles")
'
    'Disable menuitem "Delete circles"
    With ActiveDocument.CustomMenus("DeleteObjects").MenuItems("DeleteAllCircles")
        .Enabled = False
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub PieSegmentConfiguration()
'VBA487
    Dim objPieSegment As HMIPieSegment
    Set objPieSegment = ActiveDocument.HMIObjects.AddHMIObject("PieSegment1", "HMIPieSegment")
    With objPieSegment
        .StartAngle = 40
        .EndAngle = 180
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DirectConnection()
'VBA488
    Dim objButton As HMIButton
    Dim objRectangleA As HMIRectangle
    Dim objRectangleB As HMIRectangle
    Dim objEvent As HMIEvent
    Dim objDirConnection As HMIDirectConnection
    Set objRectangleA = ActiveDocument.HMIObjects.AddHMIObject("Rectangle_A", "HMIRectangle")
    Set objRectangleB = ActiveDocument.HMIObjects.AddHMIObject("Rectangle_B", "HMIRectangle")
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    With objRectangleA
        .Top = 100
        .Left = 100
    End With
    With objRectangleB
        .Top = 250
        .Left = 400
        .BackColor = RGB(255, 0, 0)
    End With
    With objButton
        .Top = 10
        .Left = 10
        .Width = 100
        .Text = "SetPosition"
    End With
'
    'Directconnection is initiated by mouseclick:
    Set objDirConnection = objButton.Events(1).Actions.AddAction(hmiActionCreationTypeDirectConnection)
    With objDirConnection
        'Sourceobject: Property "Top" of Rectangle_A
        .SourceLink.Type = hmiSourceTypeProperty
        .SourceLink.ObjectName = "Rectangle_A"
        .SourceLink.AutomationName = "Top"
'
        'Targetobject: Property "Left" of Rectangle_B
        .DestinationLink.Type = hmiDestTypeProperty
        .DestinationLink.ObjectName = "Rectangle_B"
        .DestinationLink.AutomationName = "Left"
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddActionToObjectTypeCScript()
'VBA489
    Dim objEvent As HMIEvent
    Dim objCScript As HMIScriptInfo
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_AB", "HMICircle")
'
    'C-action is initiated by click on object circle
    Set objEvent = objCircle.Events(1)
    Set objCScript = objEvent.Actions.AddAction(hmiActionCreationTypeCScript)
    MsgBox "the type of the projected event is " & objEvent.EventType
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphConfiguration()
'VBA490
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .Exponent = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SliderConfiguration()
'VBA491
    Dim objSlider As HMISlider
    Set objSlider = ActiveDocument.HMIObjects.AddHMIObject("SliderObject1", "HMISlider")
    With objSlider
        .ExtendedOperation = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ExampleForLanguageFonts()
'VBA492
    Dim colLangFonts As HMILanguageFonts
    Dim objButton As HMIButton
    Dim iStartLangID As Integer
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    iStartLangID = Application.CurrentDataLanguage
    With objButton
        .Text = "Command"
        .Width = 100
    End With
    Set colLangFonts = objButton.LDFonts
'
    'To do typesettings for french:
    With colLangFonts.ItemByLCID(1036)
        .Family = "Courier New"
        .Bold = True
        .Italic = False
        .Underlined = True
        .Size = 12
    End With
'
    'To do typesettings for english:
    With colLangFonts.ItemByLCID(1033)
        .Family = "Times New Roman"
        .Bold = False
        .Italic = True
        .Underlined = False
        .Size = 14
    End With
    With objButton
        Application.CurrentDataLanguage = 1036
        .Text = "Command"
        MsgBox "Datalanguage is changed in french"
        Application.CurrentDataLanguage = 1033
        .Text = "Command"
        MsgBox "Datalanguage is changed in english"
        Application.CurrentDataLanguage = iStartLangID
        MsgBox "Datalanguage is changed back to startlanguage."
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RectangleConfiguration()
'VBA493
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .FillColor = RGB(255, 255, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RectangleConfiguration()
'VBA494
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .Filling = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RectangleConfiguration()
'VBA495
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .Filling = True
        .FillingIndex = 50
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RectangleConfiguration()
'VBA496
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .FillStyle = 196643
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphConfiguration()
'VBA497
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .FillStyle2 = 196643
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RectangleConfiguration()
'VBA498
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .FlashBackColor = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RectangleConfiguration()
'VBA499
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .FlashBorderColor = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub StatusDisplayConfiguration()
'VBA500
    Dim objsDisplay As HMIStatusDisplay
    Set objsDisplay = ActiveDocument.HMIObjects.AddHMIObject("StatusDisplay1", "HMIStatusDisplay")
    With objsDisplay
        .FlashFlashPicture = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ButtonConfiguration()
'VBA501
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .FlashForeColor = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub StatusDisplayConfiguration()
'VBA502
    Dim objStatusDisplay As HMIStatusDisplay
    Set objStatusDisplay = ActiveDocument.HMIObjects.AddHMIObject("StatusDisplay1", "HMIStatusDisplay")
    With objStatusDisplay
        .FlashPicReferenced = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub StatusDisplayConfiguration()
'VBA503
    Dim objStatusDisplay As HMIStatusDisplay
    Set objStatusDisplay = ActiveDocument.HMIObjects.AddHMIObject("StatusDisplay1", "HMIStatusDisplay")
    With objStatusDisplay
        .FlashPicTransColor = RGB(255, 255, 0)
        .FlashPicUseTransColor = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub StatusDisplayConfiguration()
'VBA504
    Dim objStatusDisplay As HMIStatusDisplay
    Set objStatusDisplay = ActiveDocument.HMIObjects.AddHMIObject("StatusDisplay1", "HMIStatusDisplay")
    With objStatusDisplay
'
        'To use this example copy a Bitmap-Graphic
        'to the "GraCS"-Folder of the actual project.
        'Replace the picturename "Testpicture.BMP" with the name of
        'the picture you copied
        .FlashPicture = "Testpicture.BMP"
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub StatusDisplayConfiguration()
'VBA505
    Dim objStatusDisplay As HMIStatusDisplay
    Set objStatusDisplay = ActiveDocument.HMIObjects.AddHMIObject("StatusDisplay1", "HMIStatusDisplay")
    With objStatusDisplay
        .FlashPicTransColor = RGB(255, 255, 0)
        .FlashPicUseTransColor = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA506
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .FlashRate = 1
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ButtonConfiguration()
'VBA507
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .FlashRateBackColor = 1
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ButtonConfiguration()
'VBA508
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .FlashRateBorderColor = 1
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub StatusDisplayConfiguration()
'VBA509
    Dim objStatusDisplay As HMIStatusDisplay
    Set objStatusDisplay = ActiveDocument.HMIObjects.AddHMIObject("StatusDisplay1", "HMIStatusDisplay")
    With objStatusDisplay
        .FlashRateFlashPic = 1
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ButtonConfiguration()
'VBA510
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .FlashRateForeColor = 1
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowFolderItems()
'VBA511
    Dim colFolderItems As HMIFolderItems
    Dim objFolderItem As HMIFolderItem
    Dim iAnswer As Integer
    Dim iMaxFolder As Integer
    Dim iMaxSymbolLib As Integer
    Dim iSymbolLibIndex As Integer
    Dim iSubFolderIndex As Integer
    Dim strSubFolderName As String
    Dim strFolderItemName As String

    'To determine the number of symbollibraries:
    iMaxSymbolLib = Application.SymbolLibraries.Count
    iSymbolLibIndex = 1
    For iSymbolLibIndex = 1 To iMaxSymbolLib
        With Application.SymbolLibraries(iSymbolLibIndex)
            Set colFolderItems = .FolderItems
'
            'determine the number of folders in actual symbollibrary:
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



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ButtonConfiguration()
'VBA513
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .FONTBOLD = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ButtonConfiguration()
'VBA514
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .FONTITALIC = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ButtonConfiguration()
'VBA515
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .FONTNAME = "Arial"
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ButtonConfiguration()
'VBA516
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .FONTSIZE = 10
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ButtonConfiguration()
'VBA517
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .FontUnderline = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ButtonConfiguration()
'VBA518
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .ForeColor = RGB(255, 0, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ButtonConfiguration()
'VBA519
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .ForeFlashColorOff = RGB(255, 0, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ButtonConfiguration()
'VBA520
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .ForeFlashColorOff = RGB(255, 255, 255)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ActiveDocumentConfiguration()
'VBA521
    Application.ActiveDocument.Grid = True
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ActiveDocumentConfiguration()
'VBA522
    Application.ActiveDocument.Grid = True
    Application.ActiveDocument.GridColor = RGB(0, 0, 255)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ActiveDocumentConfiguration()
'VBA523
    Application.ActiveDocument.Grid = True
    Application.ActiveDocument.GridHeight = 8
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ActiveDocumentConfiguration()
'VBA524
    Application.ActiveDocument.Grid = True
    Application.ActiveDocument.GridWidth = 8
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateGroup()
'VBA526
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim objEllipseSegment As HMIEllipseSegment
    Dim objGroup As HMIGroup
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    With objCircle
        .Top = 40
        .Left = 40
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 80
        .Selected = True
    End With
    MsgBox "Objects selected!"
    Set objGroup = ActiveDocument.Selection.CreateGroup
    objGroup.ObjectName = "Group1"
    Set objEllipseSegment = ActiveDocument.HMIObjects.AddHMIObject("EllipseSegment", "HMIEllipseSegment")
'
    'Add one object to the existing group
    objGroup.GroupedHMIObjects.Add ("EllipseSegment")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ActiveDocumentConfiguration()
'VBA527
    Application.ActiveDocument.Height = 1600
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub IOFieldConfiguration()
'VBA528
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .HiddenInput = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddCircle()
'VBA529
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("my Circle", "HMICircle")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ButtonConfiguration()
'VBA530
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .Hotkey = 116
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphConfiguration()
'VBA531
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .Hysteresis = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphConfiguration()
'VBA532
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .Hysteresis = True
        .HysteresisRange = 4
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateToolbar()
'VBA533
    Dim objToolbar As HMIToolbar
    Dim objToolbarItem As HMIToolbarItem
    Dim strFileWithPath
    Set objToolbar = ActiveDocument.CustomToolbars.Add("Tool1_1")
    Set objToolbarItem = objToolbar.ToolbarItems.InsertToolbarItem(1, "ti1_1", "myFirstToolbaritem")
    Set objToolbarItem = objToolbar.ToolbarItems.InsertToolbarItem(2, "ti1_2", "mySecondToolbaritem")
'
    'ITo use this example copy a *.ICO-Graphic
    'to the "GraCS"-Folder of the actual project.
    'Replace the filename "EZSTART.ICO" in the next commandline
    'with the name of the ICO-Graphic you copied
    strFileWithPath = Application.ApplicationDataPath & "EZSTART.ICO"
'
    'To assign the symbol-icon to the first toolbaritem
    objToolbar.ToolbarItems(1).Icon = strFileWithPath
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub PolyLineCoordsOutput()
'VBA534
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



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateOptionGroup()
'VBA535
    Dim objRadioBox As HMIOptionGroup
    Dim iIndex As Integer
    Set objRadioBox = ActiveDocument.HMIObjects.AddHMIObject("RadioBox_1", "HMIOptionGroup")
    With objRadioBox
        .Height = 100
        .Width = 180
        .BoxCount = 4
        For iIndex = 1 To .BoxCount
            .index = iIndex
            .Text = "myCustomText" & .index
        Next iIndex
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowInternalNameOfFolderItem()
'VBA536
    Dim objGlobalLib As HMISymbolLibrary
    Set objGlobalLib = Application.SymbolLibraries(1)
    MsgBox objGlobalLib.FolderItems(2).Folder(2).Folder.Item(1).Name
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ActiveDocumentConfiguration()
'VBA537
    Application.ActiveDocument.Views.Add
    'If you comment out the following line
    'and recall the procedure, the output of
    'the messagebox is different
    Application.ActiveDocument.Views(1).Activate
'
    'Output state of copy:
    MsgBox Application.ActiveDocument.Views(1).IsActive
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ConnectCheck()
'VBA538
    Dim bCheck As Boolean
    Dim strStatus As String
    bCheck = Application.IsConnectedToProject
    If bCheck = True Then
        strStatus = "yes"
    Else
        strStatus = "no"
    End If
    MsgBox "Connection to project available: " & strStatus
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub Document_HMIObjectPropertyChanged(ByVal Property As IHMIProperty, CancelForwarding As Boolean)
'VBA539
    Dim objProp As HMIProperty
    Dim strStatus As String
    Set objProp = Property
'
    'Checks whether property is dynamicable
    If objProp.IsDynamicable = True Then
        strStatus = "yes"
    Else
        strStatus = "no"
    End If
    MsgBox "Property: " & objProp.Name & vbCrLf & "Value: " & objProp.value & vbCrLf & "Dynamicable: " & strStatus
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ExampleForLanguageFonts()
'VBA540
    Dim objLangFonts As HMILanguageFonts
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    objButton.Text = "Hello"
    Set objLangFonts = objButton.LDFonts
'
    'fontsettings for french:
    With objLangFonts.ItemByLCID(1036)
        .Family = "Courier New"
        .Bold = True
        .Italic = False
        .Underlined = True
        .Size = 12
    End With
'
    'fontsettings for english:
    With objLangFonts.ItemByLCID(1033)
        .Family = "Times New Roman"
        .Bold = False
        .Italic = True
        .Underlined = False
        .Size = 14
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GetHeight()
'VBA541
    Dim objGroup As HMIGroup
    'Next line uses the property "Item" to get a group by name
    Set objGroup = ActiveDocument.HMIObjects.Item("Group1")
    'Otherwise next line uses index to identify a groupobject
    MsgBox "The height of object 2 is: " & objGroup.GroupedHMIObjects.Item(2).Height
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub TextListConfiguration()
'VBA542
    Dim objTextList As HMITextList
    Set objTextList = ActiveDocument.HMIObjects.AddHMIObject("myTextList", "HMITextList")
    With objTextList
        .ItemBorderStyle = 1
        .ItemBorderBackColor = RGB(255, 0, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub TextListConfiguration()
'VBA543
    Dim objTextList As HMITextList
    Set objTextList = ActiveDocument.HMIObjects.AddHMIObject("myTextList", "HMITextList")
    With objTextList
        .ItemBorderStyle = 1
        .ItemBorderColor = RGB(255, 255, 255)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub TextListConfiguration()
'VBA544
    Dim objTextList As HMITextList
    Set objTextList = ActiveDocument.HMIObjects.AddHMIObject("myTextList", "HMITextList")
    With objTextList
        .ItemBorderStyle = 1
        .ItemBorderBackColor = RGB(255, 0, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub TextListConfiguration()
'VBA545
    Dim objTextList As HMITextList
    Set objTextList = ActiveDocument.HMIObjects.AddHMIObject("myTextList", "HMITextList")
    With objTextList
        .ItemBorderWidth = 4
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateMenuItem()
'VBA546
    Dim objMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
'
    'Add new menu "Delete objects" to menubar:
    Set objMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DeleteObjects", "Delete objects")
'
    'Adds two menuitems to menu "Delete objects"
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(1, "DeleteAllRectangles", "Delete Rectangles")
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(2, "DeleteAllCircles", "Delete Circles")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub Document_MenuItemClicked(ByVal MenuItem As IHMIMenuItem)
'VBA547
    Dim strClicked As String
    Dim objMenuItem As HMIMenuItem
    Set objMenuItem = MenuItem
'
    '"strClicked can get two values:
    '(1) "DeleteAllRectangles" and
    '(2) "DeleteAllCircles"
    strClicked = objMenuItem.Key
'
    'To analyse "strClicked" with "Select Case"
    Select Case strClicked
        Case "DeleteAllRectangles"
'
            'Instead of "MsgBox" a procedurecall (e.g. "Call <Prozedurname>") can stay here
            MsgBox "'Delete rectangle' was clicked"
        Case "DeleteAllCircles"
            MsgBox "'Delete Circles' was clicked"
    End Select
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateMenuItem()
'VBA548
    Dim objMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
    Dim iIndex As Integer
    iIndex = 1
'
    'Add new menu "Delete objects" to menubar
    Set objMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DeleteObjects", "Delete objects")
'
    'Adds two menuitems to menu "Delete objects"
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(1, "DeleteAllRectangles", "Delete rectangles")
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(2, "DeleteAllCircles", "Delete circles")
    MsgBox ActiveDocument.CustomMenus(1).Label
    For iIndex = 1 To objMenu.MenuItems.Count
        MsgBox objMenu.MenuItems(iIndex).Label
    Next iIndex
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DataLanguages()
'VBA549
    Dim colDataLang As HMIDataLanguages
    Dim objDataLang As HMIDataLanguage
    Dim nLangID As Long
    Dim strLangName As String
    Dim iAnswer As Integer
    Set colDataLang = Application.AvailableDataLanguages
    For Each objDataLang In colDataLang
        nLangID = objDataLang.LanguageID
        strLangName = objDataLang.LanguageName
        iAnswer = MsgBox(nLangID & " " & strLangName, vbOKCancel)
        If vbCancel = iAnswer Then Exit For
    Next objDataLang
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DataLanguages()
'VBA550
    Dim colDataLang As HMIDataLanguages
    Dim objDataLang As HMIDataLanguage
    Dim nLangID As Long
    Dim strLangName As String
    Dim iAnswer As Integer
    Set colDataLang = Application.AvailableDataLanguages
    For Each objDataLang In colDataLang
        nLangID = objDataLang.LanguageID
        strLangName = objDataLang.LanguageName
        iAnswer = MsgBox(nLangID & " " & strLangName, vbOKCancel)
        If vbCancel = iAnswer Then Exit For
    Next objDataLang
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub TextListConfiguration()
'VBA551
    Dim objTextList As HMITextList
    Set objTextList = ActiveDocument.HMIObjects.AddHMIObject("myTextList", "HMITextList")
    With objTextList
        .LanguageSwitch = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ActiveDocumentConfiguration()
'VBA552
    Dim varLastDocChange As Variant
    varLastDocChange = Application.ActiveDocument.LastChange
    MsgBox "Last changing: " & varLastDocChange
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RectangleConfiguration()
'VBA553
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .Layer = 4
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA554
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer00Checked = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA555
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer00Checked = True
        .Layer00Color = RGB(255, 0, 255)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA556
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer00Checked = True
        .Layer00Value = 0
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA557
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer01Checked = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA558
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer01Checked = True
        .Layer01Color = RGB(255, 0, 255)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA559
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer01Checked = True
        .Layer01Value = 10
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA560
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer02Checked = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA561
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer02Checked = True
        .Layer02Color = RGB(255, 0, 255)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA562
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer02Checked = True
        .Layer02Value = 0
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA563
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer03Checked = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA564
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer03Checked = True
        .Layer03Color = RGB(255, 0, 255)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA565
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer03Checked = True
        .Layer03Value = 30
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA566
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer04Checked = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA567
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer04Checked = True
        .Layer04Color = RGB(255, 0, 255)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA568
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer04Checked = True
        .Layer04Value = 40
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA569
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer05Checked = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA570
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer05Checked = True
        .Layer05Color = RGB(255, 0, 255)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA571
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer05Checked = True
        .Layer05Value = 50
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA572
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer06Checked = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA573
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer06Checked = True
        .Layer06Color = RGB(255, 0, 255)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA574
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer06Checked = True
        .Layer06Value = 60
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA575
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer07Checked = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA576
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer07Checked = True
        .Layer07Color = RGB(255, 0, 255)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA577
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer07Checked = True
        .Layer07Value = 70
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA578
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer08Checked = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA579
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer08Checked = True
        .Layer08Color = RGB(255, 0, 255)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA580
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer08Checked = True
        .Layer08Value = 80
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA581
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer09Checked = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA582
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer09Checked = True
        .Layer09Color = RGB(255, 0, 255)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA583
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer09Checked = True
        .Layer09Value = 90
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA584
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer10Checked = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA585
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer10Checked = True
        .Layer10Color = RGB(255, 0, 255)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA586
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer10Checked = True
        .Layer10Value = 100
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ConfigureSettingsOfLayer()
'VBA587
    Dim objLayer As HMILayer
    Set objLayer = ActiveDocument.Layers(1)
    With objLayer
        'configure "Layer 0"
        .MinZoom = 10
        .MaxZoom = 100
        .Name = "Configured with VBA"
    End With
    'define fade-in and fade-out of objects:
    With ActiveDocument
        .LayerDecluttering = True
        .ObjectSizeDecluttering = True
        .SetDeclutterObjectSize 50, 100
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub LayerInfo()
'VBA588
    Dim colLayers As HMILayers
    Dim objLayer As HMILayer
    Dim iAnswer As Integer
    Set colLayers = ActiveDocument.Layers
    For Each objLayer In colLayers
        With objLayer
            iAnswer = MsgBox("Layername: " & .Name & vbCrLf & "max. zoom:  " & .MaxZoom & vbCrLf & "min. zoom:  " & .MinZoom, vbOKCancel)
        End With
        If vbCancel = iAnswer Then Exit For
    Next objLayer
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowLanguageFont()
'VBA589
    Dim colLanguageFonts As HMILanguageFonts
    Dim objLanguageFont As HMILanguageFont
    Dim objButton As HMIButton
    Dim iMax As Integer
    Dim iAnswer As Integer
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    Set colLanguageFonts = objButton.LDFonts
    iMax = colLanguageFonts.Count
    For Each objLanguageFont In colLanguageFonts
        iAnswer = MsgBox("Projected fonts: " & iMax & vbCrLf & "Language-ID: " & objLanguageFont.LanguageID, vbOKCancel)
        If vbCancel = iAnswer Then Exit For
    Next objLanguageFont
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateMenuItem()
'VBA590
    Dim objMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
    Dim objLangText As HMILanguageText
'
    'Add new menu "Delete objects" to menubar:
    Set objMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DeleteObjects", "Delete objects")
'
    'Add two menuitems to the new menu
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(1, "DeleteAllRectangles", "Delete rectangles")
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(2, "DeleteAllCircles", "Delete circles")
'
    'Define foreign-language labels for menu "Delete objects":
    Set objLangText = objMenu.LDLabelTexts.Add(1033, "English_Delete objects")
    Set objLangText = objMenu.LDLabelTexts.Add(1032, "Greek_Delete objects")
    Set objLangText = objMenu.LDLabelTexts.Add(1034, "Spanish_Delete objects")
    Set objLangText = objMenu.LDLabelTexts.Add(1036, "French_Delete objects")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub LDLabelInfo()
'VBA591
    Dim colLangTexts As HMILanguageTexts
    Dim objLangText As HMILanguageText
    Dim iAnswer As Integer
'
    'Save all labels of menu into collection "colLangTexts":
    Set colLangTexts = ActiveDocument.CustomMenus("DeleteObjects").LDLabelTexts
    For Each objLangText In colLangTexts
        iAnswer = MsgBox(objLangText.DisplayName, vbOKCancel)
        If vbCancel = iAnswer Then Exit For
    Next objLangText
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub LDLabelInfo()
'VBA592
    Dim colLayerLngTexts As HMILanguageTexts
    Dim objLayerLngText As HMILanguageText
    Dim iIndex As Integer
    Dim iAnswer As Integer
    Dim strResult As String
    iIndex = 1
    For iIndex = 1 To ActiveDocument.Layers.Count
'
        'Save all labels of layers into collection of "colLayerLngTexts":
        Set colLayerLngTexts = ActiveDocument.Layers(iIndex).LDNames
        For Each objLayerLngText In colLayerLngTexts
            strResult = strResult & vbCrLf & objLayerLngText.LanguageID & " - " & objLayerLngText.DisplayName
        Next objLayerLngText
        iAnswer = MsgBox(strResult, vbOKCancel)
        strResult = ""
        If vbCancel = iAnswer Then Exit For
    Next iIndex
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateMenuItem()
'VBA593
    Dim objMenu As HMIMenu
    Dim objMenuItem1 As HMIMenuItem
    Dim objMenuItem2 As HMIMenuItem
    Dim objLangStateText As HMILanguageText
'
    'Add new menu "Delete objects" to menubar:
    Set objMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DeleteObjects", "Delete objects")
'
    'Add two menuitems to the new menu
    Set objMenuItem1 = objMenu.MenuItems.InsertMenuItem(1, "DeleteAllRectangles", "Delete rectangles")
    Set objMenuItem2 = objMenu.MenuItems.InsertMenuItem(2, "DeleteAllCircles", "Delete circles")
'
    'Define foreign-language labels for menuitem "Delete rectangles":
    Set objLangStateText = objMenuItem1.LDStatusTexts.Add(1033, "English_Delete rectangles")
    Set objLangStateText = objMenuItem1.LDStatusTexts.Add(1032, "Greek_Delete rectangles")
    Set objLangStateText = objMenuItem1.LDStatusTexts.Add(1034, "Spanish_Delete rectangles")
    Set objLangStateText = objMenuItem1.LDStatusTexts.Add(1036, "French_Delete rectangles")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub LDStatusTextInfo()
'VBA594
    Dim colMenuItems As HMIMenuItems
    Dim objMenuItem As HMIMenuItem
    Dim colStatusLngTexts As HMILanguageTexts
    Dim objStatusLngText As HMILanguageText
    Dim strResult As String
    Dim iAnswer As Integer
    Set colMenuItems = ActiveDocument.CustomMenus("DeleteObjects").MenuItems
    For Each objMenuItem In colMenuItems
        strResult = "Statustexts of menuitem """ & objMenuItem.Label & """"
        Set colStatusLngTexts = objMenuItem.LDStatusTexts
        For Each objStatusLngText In colStatusLngTexts
            strResult = strResult & vbCrLf & objStatusLngText.DisplayName
        Next objStatusLngText
        iAnswer = MsgBox(strResult, vbOKCancel)
        If vbCancel = iAnswer Then Exit For
    Next objMenuItem
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub LDTextInfo()
'VBA595
    Dim colLDLngTexts As HMILanguageTexts
    Dim objLDLngText As HMILanguageText
    Dim objButton As HMIButton
    Dim iAnswer As Integer
    Set objButton = ActiveDocument.HMIObjects("myButton")
    Set colLDLngTexts = objButton.LDTexts
    For Each objLDLngText In colLDLngTexts
        iAnswer = MsgBox(objLDLngText.DisplayName, vbOKCancel)
        If vbCancel = iAnswer Then Exit For
    Next objLDLngText
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateToolbar()
'VBA596
    Dim objToolbar As HMIToolbar
    Dim objToolbarItem As HMIToolbarItem
    Dim objLangText As HMILanguageText
    Dim strFileWithPath
'
    'Create toolbar with two toolbar-items:
    Set objToolbar = ActiveDocument.CustomToolbars.Add("Tool1_1")
    Set objToolbarItem = objToolbar.ToolbarItems.InsertToolbarItem(1, "ti1_1", "myFirstToolbaritem")
    Set objToolbarItem = objToolbar.ToolbarItems.InsertToolbarItem(2, "ti1_2", "mySecondToolbaritem")
'
    'In order that the example runs correct copy a *.ICO-Graphic
    'into the "GraCS"-Folder of the actual project.
    'Replace the filename "EZSTART.ICO" in the next commandline
    'with the name of the ICO-Graphic you copied
    strFileWithPath = Application.ApplicationDataPath & "EZSTART.ICO"
'
'
    'To assign the symbol-icon to the first toolbaritem
    objToolbar.ToolbarItems(1).Icon = strFileWithPath
'
    'Define foreign-language tooltiptexts
    Set objLangText = objToolbar.ToolbarItems(1).LDTooltipTexts.Add(1036, "French_Tooltiptext")
    Set objLangText = objToolbar.ToolbarItems(1).LDTooltipTexts.Add(1034, "Spanish_Tooltiptext")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub LDTooltipInfo()
'VBA597
    Dim colLangTexts As HMILanguageTexts
    Dim objLangText As HMILanguageText
    Dim iAnswer As Integer
    Set colLangTexts = ActiveDocument.CustomToolbars(1).ToolbarItems(1).LDTooltipTexts
    For Each objLangText In colLangTexts
        iAnswer = MsgBox(objLangText.DisplayName, vbOKCancel)
        If vbCancel = iAnswer Then Exit For
    Next objLangText
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RectangleConfiguration()
'VBA598
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .Left = 40
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphConfiguration()
'VBA599
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .LeftComma = 4
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA600
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .LightEffect = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA601
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeLimitHigh4 = False
        'Activate monitoring
        .CheckLimitHigh4 = True
        'Set barcolor to "red"
        .ColorLimitHigh4 = RGB(255, 0, 0)
        'Set upper limit to "70"
        .LimitHigh4 = 70
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA602
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeLimitHigh5 = False
        'Activate monitoring
        .CheckLimitHigh5 = True
        'Set barcolor to "black"
        .ColorLimitHigh5 = RGB(0, 0, 0)
        'Set upper limit to "80"
        .LimitHigh4 = 80
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA603
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeLimitLow4 = False
        'Activate monitoring
        .CheckLimitLow4 = True
        'Set barcolor to "green"
        .ColorLimitLow4 = RGB(0, 255, 0)
        'Set lower limit to "5"
        .LimitLow4 = 5
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA604
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis to absolute
        .TypeLimitLow5 = False
        'Activate monitoring
        .CheckLimitLow5 = True
        'Set barcolor to "white"
        .ColorLimitLow5 = RGB(255, 255, 255)
        'Set lower limit to "0"
        .LimitLow5 = 0
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub IOFieldConfiguration()
'VBA605
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .DataFormat = 1
        .LimitMax = 100
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub IOFieldConfiguration()
'VBA606
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .DataFormat = 1
        .LimitMin = 0
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub TextListConfiguration()
'VBA607
    Dim objTextList As HMITextList
    Set objTextList = ActiveDocument.HMIObjects.AddHMIObject("myTextList", "HMITextList")
    With objTextList
        .ListType = 0
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA608
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .LockStatus = True
        .LockBackColor = RGB(255, 0, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA609
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .LockStatus = True
        .LockBackColor = RGB(255, 0, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA610
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .LockStatus = True
        .LockText = "gesperrt"
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA611
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .LockStatus = True
        .LockTextColor = RGB(0, 255, 255)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphConfiguration()
'VBA612
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .LongStrokesBold = False
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphConfiguration()
'VBA613
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .LongStrokesOnly = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphConfiguration()
'VBA614
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .LongStrokesSize = 10
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphConfiguration()
'VBA615
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .LongStrokesTextEach = 3
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateDocumentMenusUsingMacroProperty()
'VBA616
    Dim objDocMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
    Set objDocMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DocMenu1", "Doc_Menu_1")
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(1, "dmItem1_1", "My first menuitem")
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(2, "dmItem1_2", "My second menuitem")
'
    'To assign a macro to every menuitem:
    With ActiveDocument.CustomMenus("DocMenu1")
        .MenuItems("dmItem1_1").Macro = "TestMacro1"
        .MenuItems("dmItem1_2").Macro = "TestMacro2"
    End With
End Sub


Sub TestMacro1()
    MsgBox "TestMacro1 is executed"
End Sub
 

Sub TestMacro2()
    MsgBox "TestMacro2 is executed"
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphConfiguration()
'VBA617
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .Marker = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphConfiguration()
'VBA618
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .Max = 10
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ApplicationWindowConfig()
'VBA619
    Dim objAppWindow As HMIApplicationWindow
    Set objAppWindow = ActiveDocument.HMIObjects.AddHMIObject("AppWindow1", "HMIApplicationWindow")
    With objAppWindow
        .MaximizeButton = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub LayerInfo()
'VBA620
    Dim colLayers As HMILayers
    Dim objSingleLayer As HMILayer
    Dim iAnswer As Integer
    Set colLayers = ActiveDocument.Layers
    For Each objSingleLayer In colLayers
        With objSingleLayer
            iAnswer = MsgBox("Layername: " & .Name & vbCrLf & "Min. zoom:  " & .MinZoom & vbCrLf & "Max. zoom:  " & .MaxZoom, vbOKCancel)
        End With
        If vbCancel = iAnswer Then Exit For
    Next objSingleLayer
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA621
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCGUBackColorOff = RGB(255, 0, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA622
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCGUBackColorOn = RGB(255, 255, 255)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA623
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCGUBackFlash = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA624
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCGUTextColorOff = RGB(0, 0, 255)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA625
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCGUTextColorOn = RGB(0, 0, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA626
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCGUTextFlash = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA627
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCKOBackColorOff = RGB(255, 0, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA628
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCKOBackColorOn = RGB(255, 255, 255)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA629
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCKOBackFlash = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA630
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCKOTextColorOff = RGB(0, 0, 255)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA631
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCKOTextColorOn = RGB(0, 0, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA632
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCKOTextFlash = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA633
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCKQBackColorOff = RGB(255, 0, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA634
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCKQBackColorOn = RGB(255, 255, 255)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA635
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCKQBackFlash = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA636
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCKQTextColorOff = RGB(0, 0, 255)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA637
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCKQTextColorOn = RGB(0, 0, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA638
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCKQTextFlash = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA639
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MessageClass = 0
        .MCText = "Alarm High"
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateMenuItem()
'VBA640
    Dim objMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
    Dim iIndex As Integer
    iIndex = 1
'
    'Add new menu "Delete objects" to the menubar:
    Set objMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DeleteObjects", "Delete objects")
'
    'Add two menuitems to menu "Delete objects"
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(1, "DeleteAllRectangles", "Delete rectangles")
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(2, "DeleteAllCircles", "Delete circles")
'
    'Output label of menu:
    MsgBox ActiveDocument.CustomMenus(1).Label
'
    'Output labels of all menuitems:
    For iIndex = 1 To objMenu.MenuItems.Count
        MsgBox objMenu.MenuItems(iIndex).Label
    Next iIndex
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowMenuTypes()
'VBA641
    Dim iMaxMenuItems As Integer
    Dim iMenuItemType As Integer
    Dim strMenuItemType As String
    Dim iIndex As Integer
    iMaxMenuItems = ActiveDocument.CustomMenus(1).MenuItems.Count
    For iIndex = 1 To iMaxMenuItems
        iMenuItemType = ActiveDocument.CustomMenus(1).MenuItems(iIndex).MenuItemType
        Select Case iMenuItemType
            Case 0
                strMenuItemType = "Trennstrich (Separator)"
            Case 1
                strMenuItemType = "Untermenü (SubMenu)"
            Case 2
                strMenuItemType = "Menüeintrag (MenuItem)"
        End Select
        MsgBox iIndex & ". Menuitemtype: " & strMenuItemType
    Next iIndex
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA642
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MessageClass = 0
        .MCGUBackColorOff = RGB(255, 0, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphConfiguration()
'VBA643
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .Min = 1
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub LayerInfo()
'VBA644
    Dim colLayers As HMILayers
    Dim objLayer As HMILayer
    Dim strMaxZoom As String
    Dim strMinZoom As String
    Dim strLayerName As String
    Dim iAnswer As Integer
    Set colLayers = ActiveDocument.Layers
    For Each objLayer In colLayers
        With objLayer
            strMinZoom = .MinZoom
            strMaxZoom = .MaxZoom
            strLayerName = .Name
            iAnswer = MsgBox("Layername: " & strLayerName & vbCrLf & "Min. zoom:  " & strMinZoom & vbCrLf & "Max. zoom:  " & strMaxZoom, vbOKCancel)
        End With
        If vbCancel = iAnswer Then Exit For
    Next objLayer
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CheckModificationOfActiveDocument()
'VBA645
    Dim strCheck As String
    Dim bModified As Boolean
    bModified = ActiveDocument.Modified
    Select Case bModified
        Case True
            strCheck = "Active document is modified"
        Case False
            strCheck = "Active document is not modified"
    End Select
    MsgBox strCheck
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ApplicationWindowConfig()
'VBA646
    Dim objAppWindow As HMIApplicationWindow
    Set objAppWindow = ActiveDocument.HMIObjects.AddHMIObject("AppWindow1", "HMIApplicationWindow")
    With objAppWindow
        .Moveable = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub LayerInfo()
'VBA647
    Dim colLayers As HMILayers
    Dim objLayer As HMILayer
    Dim strMaxZoom As String
    Dim strMinZoom As String
    Dim strLayerName As String
    Dim iAnswer As Integer
    Set colLayers = ActiveDocument.Layers
    For Each objLayer In colLayers
        With objLayer
            strMinZoom = .MinZoom
            strMaxZoom = .MaxZoom
            strLayerName = .Name
            iAnswer = MsgBox("Layername: " & strLayerName & vbCrLf & "Min. zoom:  " & strMinZoom & vbCrLf & "Max. zoom:  " & strMaxZoom, vbOKCancel)
        End With
        If vbCancel = iAnswer Then Exit For
    Next objLayer
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeBinary()
'VBA648
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_C", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeBool
        .BinaryResultInfo.NegativeValue = 20
        .BinaryResultInfo.PositiveValue = 40
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub TextListConfiguration()
'VBA649
    Dim objTextList As HMITextList
    Set objTextList = ActiveDocument.HMIObjects.AddHMIObject("myTextList", "HMITextList")
    With objTextList
        .NumberLines = 3
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DirectConnection()
'VBA650
    Dim objButton As HMIButton
    Dim objRectangleA As HMIRectangle
    Dim objRectangleB As HMIRectangle
    Dim objEvent As HMIEvent
    Dim objDirConnection As HMIDirectConnection
    Set objRectangleA = ActiveDocument.HMIObjects.AddHMIObject("Rectangle_A", "HMIRectangle")
    Set objRectangleB = ActiveDocument.HMIObjects.AddHMIObject("Rectangle_B", "HMIRectangle")
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    With objRectangleA
        .Top = 100
        .Left = 100
    End With
    With objRectangleB
        .Top = 250
        .Left = 400
        .BackColor = RGB(255, 0, 0)
    End With
    With objButton
        .Top = 10
        .Left = 10
        .Width = 100
        .Text = "SetPosition"
    End With
'
    'Directconnection is initiated by mouseclick:
    Set objDirConnection = objButton.Events(1).Actions.AddAction(hmiActionCreationTypeDirectConnection)
    With objDirConnection
        'Sourceobject: Property "Top" of Rectangle_A
        .SourceLink.Type = hmiSourceTypeProperty
        .SourceLink.ObjectName = "Rectangle_A"
        .SourceLink.AutomationName = "Top"
'
        'Targetobject: Property "Left" of Rectangle_B
        .DestinationLink.Type = hmiDestTypeProperty
        .DestinationLink.ObjectName = "Rectangle_B"
        .DestinationLink.AutomationName = "Left"
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ConfigureSettingsOfLayer()
'VBA651
    Dim objLayer As HMILayer
    Set objLayer = ActiveDocument.Layers(1)
    With objLayer
        'Configure "Layer 0"
        .MinZoom = 10
        .MaxZoom = 100
        .Name = "Configured with VBA"
    End With
    'Define fade-in and fade-out of objects:
    With ActiveDocument
        .LayerDecluttering = True
        .ObjectSizeDecluttering = True
        .SetDeclutterObjectSize 50, 100
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub PictureWindowConfig()
'VBA652
    Dim objPicWindow As HMIPictureWindow
    Set objPicWindow = ActiveDocument.HMIObjects.AddHMIObject("PicWindow1", "HMIPictureWindow")
    With objPicWindow
        .AdaptPicture = False
        .AdaptSize = False
        .Caption = True
        .CaptionText = "Picturewindow in runtime"
        .OffsetLeft = 5
        .OffsetTop = 10
        'Replace the picturename "Test.PDL" with the name of
        'an existing document from your "GraCS"-Folder of your active project
        .PictureName = "Test.PDL"
        .ScrollBars = True
        .ServerPrefix = ""
        .TagPrefix = "Struct."
        .UpdateCycle = 5
        .Zoom = 100
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub PictureWindowConfig()
'VBA653
    Dim objPicWindow As HMIPictureWindow
    Set objPicWindow = ActiveDocument.HMIObjects.AddHMIObject("PicWindow1", "HMIPictureWindow")
    With objPicWindow
        .AdaptPicture = False
        .AdaptSize = False
        .Caption = True
        .CaptionText = "Picturewindow in runtime"
        .OffsetLeft = 5
        .OffsetTop = 10
        'Replace the picturename "Test.PDL" with the name of
        'an existing document from your "GraCS"-Folder of your active project
        .PictureName = "Test.PDL"
        .ScrollBars = True
        .ServerPrefix = ""
        .TagPrefix = "Struct."
        .UpdateCycle = 5
        .Zoom = 100
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ApplicationWindowConfig()
'VBA654
    Dim objAppWindow As HMIApplicationWindow
    Set objAppWindow = ActiveDocument.HMIObjects.AddHMIObject("AppWindow1", "HMIApplicationWindow")
    With objAppWindow
        .OnTop = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowOperationStatusOfAllObjects()
'VBA655
    Dim objObject As HMIObject
    Dim bStatus As Boolean
    Dim strStatus As String
    Dim strName As String
    Dim iMax As Integer
    Dim iIndex As Integer
    Dim iAnswer As Integer
    iMax = ActiveDocument.HMIObjects.Count
    iIndex = 1
    For iIndex = 1 To iMax
        strName = ActiveDocument.HMIObjects(iIndex).ObjectName
        bStatus = ActiveDocument.HMIObjects(iIndex).Operation
        Select Case bStatus
            Case True
                strStatus = "yes"
            Case False
                strStatus = "no"
        End Select
        iAnswer = MsgBox("Object: " & strName & vbCrLf & "Operator-Control enable: " & strStatus, vbOKCancel)
        If vbCancel = iAnswer Then Exit For
    Next iIndex
    If 0 = iMax Then MsgBox "No objects in the active document."
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub IOFieldConfiguration()
'VBA656
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .OperationReport = True
        .OperationMessage = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub IOFieldConfiguration()
'VBA657
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .OperationReport = True
        .OperationMessage = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ButtonConfiguration()
'VBA658
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .Width = 150
        .Height = 150
        .Text = "Text is displayed vertical"
        .Orientation = False
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub IOFieldConfiguration()
'VBA659
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .DataFormat = 1
        .OutputFormat = "99,999"
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub IOFieldConfiguration()
'VBA660
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .OutputValue = "00"
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ExampleForParent()
'VBA661
    Dim objView As HMIView
    Set objView = ActiveDocument.Views.Add
    MsgBox objView.Parent.Name
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowDocumentPath()
'VBA663
    MsgBox ActiveDocument.Path
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowInternalNameOfFolderItem()
'VBA664
    Dim objGlobalLib As HMISymbolLibrary
    Set objGlobalLib = Application.SymbolLibraries(1)
    MsgBox objGlobalLib.FolderItems(2).Folder(2).Folder.Item(1).PathName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RoundButtonConfiguration()
'VBA665
    Dim objRoundButton As HMIRoundButton
    Set objRoundButton = ActiveDocument.HMIObjects.AddHMIObject("RButton1", "HMIRoundButton")
    With objRoundButton
        .PicDeactReferenced = False
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RoundButtonConfiguration()
'VBA666
    Dim objRoundButton As HMIRoundButton
    Set objRoundButton = ActiveDocument.HMIObjects.AddHMIObject("RButton1", "HMIRoundButton")
    With objRoundButton
        .PicDeactTransparent = RGB(255, 0, 0)
        .PicDeactUseTransColor = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RoundButtonConfiguration()
'VBA667
    Dim objRoundButton As HMIRoundButton
    Set objRoundButton = ActiveDocument.HMIObjects.AddHMIObject("RButton1", "HMIRoundButton")
    With objRoundButton
        .PicDeactTransparent = RGB(255, 0, 0)
        .PicDeactUseTransColor = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RoundButtonConfiguration()
'VBA668
    Dim objRoundButton As HMIRoundButton
    Set objRoundButton = ActiveDocument.HMIObjects.AddHMIObject("RButton1", "HMIRoundButton")
    With objRoundButton
        .PicDownReferenced = False
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RoundButtonConfiguration()
'VBA669
    Dim objRoundButton As HMIRoundButton
    Set objRoundButton = ActiveDocument.HMIObjects.AddHMIObject("RButton1", "HMIRoundButton")
    With objRoundButton
        .PicDownTransparent = RGB(255, 255, 0)
        .PicDownUseTransColor = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RoundButtonConfiguration()
'VBA670
    Dim objRoundButton As HMIRoundButton
    Set objRoundButton = ActiveDocument.HMIObjects.AddHMIObject("RButton1", "HMIRoundButton")
    With objRoundButton
        .PicDownTransparent = RGB(255, 255, 0)
        .PicDownUseTransColor = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GraphicObjectConfiguration()
'VBA671
    Dim objGraphicObject As HMIGraphicObject
    Set objGraphicObject = ActiveDocument.HMIObjects.AddHMIObject("GraphicObject1", "HMIGraphicObject")
    With objGraphicObject
        .PicReferenced = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GraphicObjectConfiguration()
'VBA672
    Dim objGraphicObject As HMIGraphicObject
    Set objGraphicObject = ActiveDocument.HMIObjects.AddHMIObject("GraphicObject1", "HMIGraphicObject")
    With objGraphicObject
        .PicTransColor = 16711680
        .PicUseTransColor = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ButtonConfiguration()
'VBA673
    Dim objRoundButton As HMIRoundButton
    Set objRoundButton = ActiveDocument.HMIObjects.AddHMIObject("RButton1", "HMIRoundButton")
    With objRoundButton
'
        'Toi use this example copy a Bitmap-Graphic
        'to the "GraCS"-Folder of the actual project.
        'Replace the picturename "TestPicture1.BMP" with the name of
        'the picture you copied
        .PictureDeactivated = "TestPicture1.BMP"
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ButtonConfiguration()
'VBA674
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
    '
        'To use this example copy two Bitmap-Graphics
        'to the "GraCS"-Folder of the actual project.
        'Replace the picturenames "TestPicture1.BMP" and "TestPicture2.BMP"
        'with the names of the pictures you copied
        .PictureDown = "TestPicture1.BMP"
        .PictureUp = "TestPicture2.BMP"
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub PictureWindowConfig()
'VBA675
    Dim objPicWindow As HMIPictureWindow
    Set objPicWindow = ActiveDocument.HMIObjects.AddHMIObject("PicWindow1", "HMIPictureWindow")
    With objPicWindow
        .AdaptPicture = False
        .AdaptSize = False
        .Caption = True
        .CaptionText = "Picturewindow in runtime"
        .OffsetLeft = 5
        .OffsetTop = 10
        'Replace the picturename "Test.PDL" with the name of
        'an existing document from your "GraCS"-Folder of your active project
        .PictureName = "Test.PDL"
        .ScrollBars = True
        .ServerPrefix = ""
        .TagPrefix = "Struct."
        .UpdateCycle = 5
        .Zoom = 100
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ButtonConfiguration()
'VBA676
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
    '
        'To use this example copy two Bitmap-Graphics
        'to the "GraCS"-Folder of the actual project.
        'Replace the picturenames "TestPicture1.BMP" and "TestPicture2.BMP"
        'with the names of the pictures you copied
        .PictureDown = "TestPicture1.BMP"
        .PictureUp = "TestPicture2.BMP"
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RoundButtonConfiguration()
'VBA677
    Dim objRoundButton As HMIRoundButton
    Set objRoundButton = ActiveDocument.HMIObjects.AddHMIObject("RButton1", "HMIRoundButton")
    With objRoundButton
        .PicUpReferenced = False
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RoundButtonConfiguration()
'VBA678
    Dim objRoundButton As HMIRoundButton
    Set objRoundButton = ActiveDocument.HMIObjects.AddHMIObject("RButton1", "HMIRoundButton")
    With objRoundButton
        .PicUpTransparent = RGB(0, 0, 255)
        .PicUpUseTransColor = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RoundButtonConfiguration()
'VBA679
    Dim objRoundButton As HMIRoundButton
    Set objRoundButton = ActiveDocument.HMIObjects.AddHMIObject("RButton1", "HMIRoundButton")
    With objRoundButton
        .PicUpTransparent = RGB(0, 0, 255)
        .PicUpUseTransColor = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GraphicObjectConfiguration()
'VBA680
    Dim objGraphicObject As HMIGraphicObject
    Set objGraphicObject = ActiveDocument.HMIObjects.AddHMIObject("GraphicObject1", "HMIGraphicObject")
    With objGraphicObject
        .PicTransColor = RGB(0, 0, 255)
        .PicUseTransColor = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub LineAdd()
'VBA682
    Dim objLine As HMILine
    Dim objEvent As HMIEvent
    Set objLine = ActiveDocument.HMIObjects.AddHMIObject("myLine", "HMILine")
    With objLine
        .BorderColor = RGB(255, 0, 0)
        .index = hmiLineIndexTypeStartPoint
        .ActualPointLeft = 12
        .ActualPointTop = 34
        .index = hmiLineIndexTypeEndPoint
        .ActualPointLeft = 74
        .ActualPointTop = 64
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowPositionOfCustomMenuItems()
'VBA683
    Dim objMenu As HMIMenu
    Dim iMaxMenuItems As Integer
    Dim iPosition As Integer
    Dim iIndex As Integer
    Set objMenu = ActiveDocument.CustomMenus(1)
    iMaxMenuItems = objMenu.MenuItems.Count
    For iIndex = 1 To iMaxMenuItems
        iPosition = objMenu.MenuItems(iIndex).Position
        MsgBox "Position of the " & iIndex & ". menuitem: " & iPosition
    Next iIndex
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeBool()
'VBA684
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_C", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeBool
        .BinaryResultInfo.NegativeValue = 20
        .BinaryResultInfo.PositiveValue = 40
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA685
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        'Depth-angle a = 15 degrees
        .AngleAlpha = 15
        .PredefinedAngles = 1
        'Depth-angle b = 45 degrees
        .AngleBeta = 45
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RoundButtonConfiguration()
'VBA686
    Dim objRoundButton As HMIRoundButton
    Set objRoundButton = ActiveDocument.HMIObjects.AddHMIObject("RButton1", "HMIRoundButton")
    With objRoundButton
        .Pressed = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HMI3DBarGraphConfiguration()
'VBA687
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        'Depth-angle a = 15 degrees
        .AngleAlpha = 15
        'Depth-angle b = 45 degrees
        .AngleBeta = 45
        .Process = 100
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowProfileName()
'VBA688
    MsgBox Application.ProfileName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddActiveXControl()
'VBA689
    Dim objActiveXControl As HMIActiveXControl
    Set objActiveXControl = ActiveDocument.HMIObjects.AddActiveXControl("WinCC_Gauge", "XGAUGE.XGaugeCtrl.1")
    With ActiveDocument
        .HMIObjects("WinCC_Gauge").Top = 40
        .HMIObjects("WinCC_Gauge").Left = 40
        MsgBox "ProgID of ActiveX-control: " & .HMIObjects("WinCC_Gauge").ProgID
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowProjectInfo()
'VBA690
    Dim iProjectType As Integer
    Dim strProjectName As String
    Dim strProjectType As String
    iProjectType = Application.ProjectType
    strProjectName = Application.ProjectName
    Select Case iProjectType
        Case 0
            strProjectType = "Single-User System"
        Case 1
            strProjectType = "Multi-User System"
        Case 2
            strProjectType = "Client System"
    End Select
    MsgBox "Projecttype: " & strProjectType & vbCrLf & "Projectname: " & strProjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowProjectInfo()
'VBA691
    Dim iProjectType As Integer
    Dim strProjectName As String
    Dim strProjectType As String
    iProjectType = Application.ProjectType
    strProjectName = Application.ProjectName
    Select Case iProjectType
        Case 0
            strProjectType = "Single-User System"
        Case 1
            strProjectType = "Multi-User System"
        Case 2
            strProjectType = "Client System"
    End Select
    MsgBox "Projecttype: " & strProjectType & vbCrLf & "Projectname: " & strProjectName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ExampleForPrototype()
'VBA692
    Dim objButton As HMIButton
    Dim objCircleA As HMICircle
    Dim objEvent As HMIEvent
    Dim objVBScript As HMIScriptInfo
    Set objCircleA = ActiveDocument.HMIObjects.AddHMIObject("CircleA", "HMICircle")
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    With objCircleA
        .Top = 100
        .Left = 100
    End With
    With objButton
        .Top = 10
        .Left = 10
        .Width = 200
        .Text = "Increase Radius"
    End With
    'On every mouseclick the radius have to increase:
    Set objEvent = objButton.Events(1)
    Set objVBScript = objButton.Events(1).Actions.AddAction(hmiActionCreationTypeVBScript)
    MsgBox objVBScript.Prototype
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub PieSegmentConfiguration()
'VBA693
    Dim objPieSegment As HMIPieSegment
    Set objPieSegment = ActiveDocument.HMIObjects.AddHMIObject("PieSegment1", "HMIPieSegment")
    With objPieSegment
        .StartAngle = 40
        .EndAngle = 180
        .Radius = 80
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EllipseConfiguration()
'VBA694
    Dim objEllipse As HMIEllipse
    Set objEllipse = ActiveDocument.HMIObjects.AddHMIObject("Ellipse1", "HMIEllipse")
    With objEllipse
        .RadiusHeight = 60
        .RadiusWidth = 40
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EllipseConfiguration()
'VBA695
    Dim objEllipse As HMIEllipse
    Set objEllipse = ActiveDocument.HMIObjects.AddHMIObject("Ellipse1", "HMIEllipse")
    With objEllipse
        .RadiusHeight = 60
        .RadiusWidth = 40
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub PolyLineConfiguration()
'VBA696
    Dim objPolyLine As HMIPolyLine
    Set objPolyLine = ActiveDocument.HMIObjects.AddHMIObject("PolyLine1", "HMIPolyLine")
    With objPolyLine
        .ReferenceRotationLeft = 50
        .ReferenceRotationTop = 50
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub PolyLineConfiguration()
'VBA697
    Dim objPolyLine As HMIPolyLine
    Set objPolyLine = ActiveDocument.HMIObjects.AddHMIObject("PolyLine1", "HMIPolyLine")
    With objPolyLine
        .ReferenceRotationLeft = 50
        .ReferenceRotationTop = 50
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA698
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .Relevant = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeBinary()
'VBA699
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_C", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeBool
        .BinaryResultInfo.NegativeValue = 20
        .BinaryResultInfo.PositiveValue = 40
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphConfiguration()
'VBA700
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .RightComma = 4
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub PolyLineConfiguration()
'VBA701
    Dim objPolyLine As HMIPolyLine
    Set objPolyLine = ActiveDocument.HMIObjects.AddHMIObject("PolyLine1", "HMIPolyLine")
    With objPolyLine
        .ReferenceRotationLeft = 50
        .ReferenceRotationTop = 50
        .RotationAngle = 45
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RoundRectangleConfiguration()
'VBA702
    Dim objRoundRectangle As HMIRoundRectangle
    Set objRoundRectangle = ActiveDocument.HMIObjects.AddHMIObject("RoundRectangle1", "HMIRoundRectangle")
    With objRoundRectangle
        .RoundCornerHeight = 25
        .RoundCornerWidth = 50
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RoundRectangleConfiguration()
'VBA703
    Dim objRoundRectangle As HMIRoundRectangle
    Set objRoundRectangle = ActiveDocument.HMIObjects.AddHMIObject("RoundRectangle1", "HMIRoundRectangle")
    With objRoundRectangle
        .RoundCornerHeight = 25
        .RoundCornerWidth = 50
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA704
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .SameSize = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphConfiguration()
'VBA705
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .Scaling = True
        .ScaleColor = RGB(255, 0, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphConfiguration()
'VBA706
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .Scaling = True
        .ScaleTicks = 10
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphConfiguration()
'VBA707
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .Scaling = True
        .ScaleColor = RGB(255, 0, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphConfiguration()
'VBA708
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .ScalingType = 0
        .Scaling = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ExampleForPrototype()
'VBA709
    Dim objButton As HMIButton
    Dim objCircleA As HMICircle
    Dim objEvent As HMIEvent
    Dim objVBScript As HMIScriptInfo
    Dim strScriptType As String
    Set objCircleA = ActiveDocument.HMIObjects.AddHMIObject("CircleA", "HMICircle")
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    With objCircleA
        .Top = 100
        .Left = 100
    End With
    With objButton
        .Top = 10
        .Left = 10
        .Width = 200
        .Text = "Increase Radius"
    End With
    'On every mouseclick the radius have to increase:
    Set objEvent = objButton.Events(1)
    Set objVBScript = objButton.Events(1).Actions.AddAction(hmiActionCreationTypeVBScript)
    Select Case objVBScript.ScriptType
        Case 0
            strScriptType = "VB-Skript is used"
        Case 1
            strScriptType = "C-Skript is used"
    End Select
    MsgBox strScriptType
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub PictureWindowConfig()
'VBA710
    Dim objPicWindow As HMIPictureWindow
    Set objPicWindow = ActiveDocument.HMIObjects.AddHMIObject("PicWindow1", "HMIPictureWindow")
    With objPicWindow
        .AdaptPicture = False
        .AdaptSize = False
        .Caption = True
        .CaptionText = "Picturewindow in runtime"
        .OffsetLeft = 5
        .OffsetTop = 10
        'Replace the picturename "Test.PDL" with the name of
        'an existing document from your "GraCS"-Folder of your active project
        .PictureName = "Test.PDL"
        .ScrollBars = True
        .ServerPrefix = ""
        .TagPrefix = "Struct."
        .UpdateCycle = 5
        .Zoom = 100
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateViewAndActivateView()
'VBA711
    Dim objView As HMIView
    Set objView = ActiveDocument.Views.Add
    objView.Activate
    objView.ScrollPosX = 40
    objView.ScrollPosY = 10
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateViewAndActivateView()
'VBA712
    Dim objView As HMIView
    Set objView = ActiveDocument.Views.Add
    objView.Activate
    objView.ScrollPosX = 40
    objView.ScrollPosY = 10
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub TextListConfiguration()
'VBA713
    Dim objTextList As HMITextList
    Set objTextList = ActiveDocument.HMIObjects.AddHMIObject("myTextList", "HMITextList")
    With objTextList
        .SelBGColor = RGB(255, 0, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SelectObjects()
'VBA714
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim objGroup As HMIGroup
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    With objCircle
        .Top = 40
        .Left = 40
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 80
        .Selected = True
    End With
    MsgBox "Objects selected!"
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SelectAllObjectsInActiveDocument()
'VBA715
    ActiveDocument.Selection.SelectAll
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub TextListConfiguration()
'VBA716
    Dim objTextList As HMITextList
    Set objTextList = ActiveDocument.HMIObjects.AddHMIObject("myTextList", "HMITextList")
    With objTextList
        .SelTextColor = RGB(255, 255, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddActiveXControl()
'VBA717
    Dim objActiveXControl As HMIActiveXControl
    Set objActiveXControl = ActiveDocument.HMIObjects.AddActiveXControl("WinCC_Gauge", "XGAUGE.XGaugeCtrl.1")
    With objActiveXControl
        .Top = 40
        .Left = 60
        MsgBox .Properties("ServerName").value
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub PictureWindowConfig()
'VBA718
    Dim objPicWindow As HMIPictureWindow
    Set objPicWindow = ActiveDocument.HMIObjects.AddHMIObject("PicWindow1", "HMIPictureWindow")
    With objPicWindow
        .AdaptPicture = False
        .AdaptSize = False
        .Caption = True
        .CaptionText = "Picturewindow in runtime"
        .OffsetLeft = 5
        .OffsetTop = 10
        'Replace the picturename "Test.PDL" with the name of
        'an existing document from your "GraCS"-Folder of your active project
        .PictureName = "Test.PDL"
        .ScrollBars = True
        .ServerPrefix = "my_Server::"
        .TagPrefix = "Struct."
        .UpdateCycle = 5
        .Zoom = 100
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateDocumentMenus()
'VBA719
    Dim objDocMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
    Dim objSubMenu As HMIMenuItem
'
    'Add menu to menubar:
    Set objDocMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DocMenu1", "Doc_Menu_1")
    '
    'Add menuitems to the new menu:
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(1, "dmItem1_1", "&My first Menuitem")
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(2, "dmItem1_2", "My second Menuitem")
'
    'Add seperator to menu:
    Set objMenuItem = objDocMenu.MenuItems.InsertSeparator(3, "dSeparator1_3")
'
    'Add submenu to the menu:
    Set objSubMenu = objDocMenu.MenuItems.InsertSubMenu(4, "dSubMenu1_4", "My first submenu")
'
    'Add menuitems to the submenu:
    Set objMenuItem = objSubMenu.SubMenu.InsertMenuItem(5, "dmItem1_5", "My first submenuitem")
    Set objMenuItem = objSubMenu.SubMenu.InsertMenuItem(6, "dmItem1_6", "My second submenuitem")
'
    ActiveDocument.CustomMenus("DocMenu1").MenuItems(1).ShortCut = "STRG+SHIFT+M"
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ExampleForLanguageFonts()
'VBA721
    Dim colLangFonts As HMILanguageFonts
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    objButton.Text = "DefText"
    Set colLangFonts = objButton.LDFonts
'
    'Set font-properties for french:
    With colLangFonts.ItemByLCID(1036)
        .Family = "Courier New"
        .Bold = True
        .Italic = False
        .Underlined = True
        .Size = 12
    End With
'
    'Set font-properties for english:
    With colLangFonts.ItemByLCID(1033)
        .Family = "Times New Roman"
        .Bold = False
        .Italic = True
        .Underlined = False
        .Size = 14
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ApplicationWindowConfig()
'VBA722
    Dim objAppWindow As HMIApplicationWindow
    Set objAppWindow = ActiveDocument.HMIObjects.AddHMIObject("AppWindow1", "HMIApplicationWindow")
    With objAppWindow
        .Sizeable = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SliderConfiguration()
'VBA723
    Dim objSlider As HMISlider
    Set objSlider = ActiveDocument.HMIObjects.AddHMIObject("SliderObject1", "HMISlider")
    With objSlider
        .SmallChange = 4
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ActivateSnapToGrid()
'VBA724
    ActiveDocument.SnapToGrid = True
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DirectConnection()
'VBA725
    Dim objButton As HMIButton
    Dim objRectangleA As HMIRectangle
    Dim objRectangleB As HMIRectangle
    Dim objEvent As HMIEvent
    Dim objDirConnection As HMIDirectConnection
    Set objRectangleA = ActiveDocument.HMIObjects.AddHMIObject("Rectangle_A", "HMIRectangle")
    Set objRectangleB = ActiveDocument.HMIObjects.AddHMIObject("Rectangle_B", "HMIRectangle")
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    With objRectangleA
        .Top = 100
        .Left = 100
    End With
    With objRectangleB
        .Top = 250
        .Left = 400
        .BackColor = RGB(255, 0, 0)
    End With
    With objButton
        .Top = 10
        .Left = 10
        .Width = 100
        .Text = "SetPosition"
    End With
'
    'Directconnection is initiated by mouseclick:
    Set objDirConnection = objButton.Events(1).Actions.AddAction(hmiActionCreationTypeDirectConnection)
    With objDirConnection
        'Sourceobject: Property "Top" of Rectangle_A
        .SourceLink.Type = hmiSourceTypeProperty
        .SourceLink.ObjectName = "Rectangle_A"
        .SourceLink.AutomationName = "Top"
'
        'Targetobject: Property "Left" of Rectangle_B
        .DestinationLink.Type = hmiDestTypeProperty
        .DestinationLink.ObjectName = "Rectangle_B"
        .DestinationLink.AutomationName = "Left"
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub IncreaseCircleRadiusWithVBScript()
'VBA726
    Dim objButton As HMIButton
    Dim objCircleA As HMICircle
    Dim objEvent As HMIEvent
    Dim objVBScript As HMIScriptInfo
    Dim strCode As String
    strCode = "Dim objCircle" & vbCrLf & "Set objCircle = "
    strCode = strCode & "hmiRuntime.ActiveScreen.ScreenItems(""CircleVB"")"
    strCode = strCode & vbCrLf & "objCircle.Radius = objCircle.Radius + 5"
    Set objCircleA = ActiveDocument.HMIObjects.AddHMIObject("CircleVB", "HMICircle")
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    With objCircleA
        .Top = 100
        .Left = 100
    End With
    With objButton
        .Top = 10
        .Left = 10
        .Width = 200
        .Text = "Increase Radius"
    End With
'
    'On every mouseclick the radius have to increase:
    Set objEvent = objButton.Events(1)
    Set objVBScript = objButton.Events(1).Actions.AddAction(hmiActionCreationTypeVBScript)
    objVBScript.SourceCode = strCode
    Select Case objVBScript.Compiled
        Case True
            MsgBox "Compilation ok!"
        Case False
            MsgBox "Error on compilation!"
    End Select
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub PieSegmentConfiguration()
'VBA727
    Dim PieSegment As HMIPieSegment
    Set PieSegment = ActiveDocument.HMIObjects.AddHMIObject("PieSegment1", "HMIPieSegment")
    With PieSegment
        .StartAngle = 40
        .EndAngle = 180
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateDocumentMenus()
'VBA728
    Dim objDocMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
    Dim objSubMenu As HMIMenuItem
'
    'Add menu:
    Set objDocMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DocMenu1", "Doc_Menu_1")
'
    'Add menuitems to custom-menu:
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(1, "dmItem1_1", "My first menuitem")
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(2, "dmItem1_2", "My second menuitem")
'
    'Add seperator to custom-menu:
    Set objMenuItem = objDocMenu.MenuItems.InsertSeparator(3, "dSeparator1_3")
'
    'Add submenu to custom-menu:
    Set objSubMenu = objDocMenu.MenuItems.InsertSubMenu(4, "dSubMenu1_4", "My first submenu")
'
    'Add menuitems to submenu:
    Set objMenuItem = objSubMenu.SubMenu.InsertMenuItem(5, "dmItem1_5", "My first submenuitem")
    Set objMenuItem = objSubMenu.SubMenu.InsertMenuItem(6, "dmItem1_6", "My second submenuitem")
'
    'Assign statustexts to every menuitem
    With objDocMenu
        .MenuItems(1).StatusText = "My first menuitem"
        .MenuItems(2).StatusText = "My second menuitem"
        .MenuItems(4).SubMenu.Item(1).StatusText = "My first submenuitem"
        .MenuItems(4).SubMenu.Item(2).StatusText = "My second submenuitem"
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateDocumentMenus()
'VBA730
    Dim objDocMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
    Dim objSubMenu As HMIMenuItem
'
    'Add menu:
    Set objDocMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DocMenu1", "Doc_Menu_1")
'
    'Add menuitems to custom-menu:
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(1, "dmItem1_1", "My first menuitem")
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(2, "dmItem1_2", "My second menuitem")
'
    'Add seperator to custom-menu:
    Set objMenuItem = objDocMenu.MenuItems.InsertSeparator(3, "dSeparator1_3")
'
    'Add submenu to custom-menu:
    Set objSubMenu = objDocMenu.MenuItems.InsertSubMenu(4, "dSubMenu1_4", "My first submenu")
'
    'Add menuitems to submenu:
    Set objMenuItem = objSubMenu.SubMenu.InsertMenuItem(5, "dmItem1_5", "My first submenuitem")
    Set objMenuItem = objSubMenu.SubMenu.InsertMenuItem(6, "dmItem1_6", "My second submenuitem")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowSymbolLibraries()
'VBA731
    Dim colSymbolLibraries As HMISymbolLibraries
    Dim objSymbolLibrary As HMISymbolLibrary
    Set colSymbolLibraries = Application.SymbolLibraries
    For Each objSymbolLibrary In colSymbolLibraries
        MsgBox objSymbolLibrary.Name
    Next objSymbolLibrary
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub IOFieldConfig()
'VBA732
    Dim objIOField1 As HMIIOField
    Dim objIOField2 As HMIIOField
    Set objIOField1 = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    Set objIOField2 = ActiveDocument.HMIObjects.AddHMIObject("IOField2", "HMIIOField")
    With objIOField1
        .Top = 10
        .Left = 10
        .TabOrderSwitch = 1
    End With
    With objIOField2
        .Top = 100
        .Left = 10
        .TabOrderSwitch = 2
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ConfigureTabOrder()
'VBA733
    With ActiveDocument
        .TABOrderAllHMIObjects = True
        .TABOrderKeyboard = False
        .TABOrderMouse = False
        .TABOrderOtherAction = False
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub IOFieldConfig()
'VBA734
    Dim objIOField1 As HMIIOField
    Dim objIOField2 As HMIIOField
    Set objIOField1 = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    Set objIOField2 = ActiveDocument.HMIObjects.AddHMIObject("IOField2", "HMIIOField")
    With objIOField1
        .Top = 10
        .Left = 10
        .TabOrderAlpha = 1
    End With
    With objIOField2
        .Top = 100
        .Left = 10
        .TabOrderAlpha = 2
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ConfigureTabOrder()
'VBA735
    With ActiveDocument
        .TABOrderAllHMIObjects = True
        .TABOrderKeyboard = False
        .TABOrderMouse = False
        .TABOrderOtherAction = False
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ConfigureTabOrder()
'VBA736
    With ActiveDocument
        .TABOrderAllHMIObjects = True
        .TABOrderKeyboard = False
        .TABOrderMouse = False
        .TABOrderOtherAction = False
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ConfigureTabOrder()
'VBA737
    With ActiveDocument
        .TABOrderAllHMIObjects = True
        .TABOrderKeyboard = False
        .TABOrderMouse = False
        .TABOrderOtherAction = False
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateDocumentMenus()
'VBA738
    Dim objDocMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
    Dim objSubMenu As HMIMenuItem
'
    'Add menu:
    Set objDocMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DocMenu1", "Doc_Menu_1")
'
    'Add menuitems to custom-menu:
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(1, "dmItem1_1", "My first menuitem")
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(2, "dmItem1_2", "My second menuitem")
'
    'Add seperator to custom-menu:
    Set objMenuItem = objDocMenu.MenuItems.InsertSeparator(3, "dSeparator1_3")
'
    'Add submenu to custom-menu:
    Set objSubMenu = objDocMenu.MenuItems.InsertSubMenu(4, "dSubMenu1_4", "My first submenu")
'
    'Add menuitems to submenu:
    Set objMenuItem = objSubMenu.SubMenu.InsertMenuItem(5, "dmItem1_5", "My first submenuitem")
    Set objMenuItem = objSubMenu.SubMenu.InsertMenuItem(6, "dmItem1_6", "My second submenuitem")
'
    'To place an additional information:
    With objDocMenu
        .MenuItems(1).Tag = "This is the first menuitem"
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub PictureWindowConfig()
'VBA739
    Dim objPicWindow As HMIPictureWindow
    Set objPicWindow = ActiveDocument.HMIObjects.AddHMIObject("PicWindow1", "HMIPictureWindow")
    With objPicWindow
        .AdaptPicture = False
        .AdaptSize = False
        .Caption = True
        .CaptionText = "Picturewindow in runtime"
        .OffsetLeft = 5
        .OffsetTop = 10
        'Replace the picturename "Test.PDL" with the name of
        'an existing document from your "GraCS"-Folder of your active project
        .PictureName = "Test.PDL"
        .ScrollBars = True
        .ServerPrefix = "my_Server::"
        .TagPrefix = "Struct."
        .UpdateCycle = 5
        .Zoom = 100
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ButtonConfiguration()
'VBA740
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .Text = "Button1"
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RoundButtonConfiguration()
'VBA741
    Dim objRoundButton As HMIRoundButton
    Set objRoundButton = ActiveDocument.HMIObjects.AddHMIObject("RButton1", "HMIRoundButton")
    With objRoundButton
        .Toggle = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA742
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis = absolute
        .TypeToleranceHigh = False
        'Activate monitoring
        .CheckToleranceHigh = True
        'Set barcolor = "yellow"
        .ColorToleranceHigh = RGB(255, 255, 0)
        'Set upper limit to "40"
        .ToleranceHigh = 40
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA743
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis = absolute
        .TypeToleranceLow = False
        'Activate monitoring
        .CheckToleranceLow = True
        'Set barcolor = "red"
        .ColorToleranceLow = RGB(255, 0, 0)
        'Set lower limit to "40"
        .ToleranceLow = 40
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDocumentSpecificCustomToolbar()
'VBA744
    Dim objToolbar As HMIToolbar
    Dim objToolbarItem As HMIToolbarItem
    Set objToolbar = ActiveDocument.CustomToolbars.Add("DocToolbar")
'
    'Add symbol-icon to userdefined toolbar
    Set objToolbarItem = objToolbar.ToolbarItems.InsertToolbarItem(1, "tItem1_1", "My first symbol-icon")
    Set objToolbarItem = objToolbar.ToolbarItems.InsertToolbarItem(3, "tItem1_3", "My second symbol-icon")
    Set objToolbarItem = objToolbar.ToolbarItems.InsertSeparator(2, "tSeparator1_2")
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RectangleConfiguration()
'VBA745
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .ToolTipText = "This is a rectangle"
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RectangleConfiguration()
'VBA746
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .Left = 10
        .Top = 40
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphConfiguration()
'VBA747
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .trend = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphConfiguration()
'VBA748
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .trend = True
        .TrendColor = RGB(255, 0, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicAsCSkriptToProperty()
'VBA749
    Dim objVBScript As HMIScriptInfo
    Dim objCircle As HMICircle
 
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("myCircle", "HMICircle")
    Set objVBScript = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeVBScript)
    With objVBScript
        .Trigger.Type = hmiTriggerTypeStandardCycle
        .Trigger.CycleType = hmiCycleType_2s
        .Trigger.Name = "Trigger1"
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RectangleConfiguration()
'VBA750
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        MsgBox "Objecttype: " & .Type
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA751
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis = absolute
        .TypeAlarmHigh = False
        'Activate monitoring
        .CheckAlarmHigh = True
        'Set barcolor = "yellow"
        .ColorAlarmHigh = RGB(255, 255, 0)
        'Set upper limit = "50"
        .AlarmHigh = 50
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA752
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis = absolute
        .TypeAlarmLow = False
        'Activate monitoring
        .CheckAlarmLow = True
        'Set barcolor = "yellow"
        .ColorAlarmLow = RGB(255, 255, 0)
        'Set lower limit = "10"
        .AlarmLow = 10
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA753
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis = absolute
        .TypeLimitHigh4 = False
        'Activate monitoring
        .CheckLimitHigh4 = True
        'Set barcolor = "red"
        .ColorLimitHigh4 = RGB(255, 0, 0)
        'Set upper limit = "70"
        .LimitHigh4 = 70
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA754
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis = absolute
        .TypeLimitHigh5 = False
        'Activate monitoring
        .CheckLimitHigh5 = True
        'Set barcolor = "black"
        .ColorLimitHigh5 = RGB(0, 0, 0)
        'Set upper limit = "70"
        .LimitHigh5 = 70
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA755
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis = absolute
        .TypeLimitLow4 = False
        'Activate monitoring
        .CheckLimitLow4 = True
        'Set barcolor = "green"
        .ColorLimitLow4 = RGB(0, 255, 0)
        'Set lower limit = "5"
        .LimitLow4 = 5
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA756
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis = absolute
        .TypeLimitLow5 = False
        'Activate monitoring
        .CheckLimitLow5 = True
        'Set barcolor = "white"
        .ColorLimitLow5 = RGB(255, 255, 255)
        'Set lower limit = "0"
        .LimitLow5 = 0
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA757
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis = absolute
        .TypeToleranceHigh = False
        'Activate monitoring
        .CheckToleranceHigh = True
        'Set barcolor = "yellow"
        .ColorToleranceHigh = RGB(255, 255, 0)
        'Set upper limit = "40"
        .ToleranceHigh = 40
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA758
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis = absolute
        .TypeToleranceLow = False
        'Activate monitoring
        .CheckToleranceLow = True
        'Set barcolor = "red"
        .ColorToleranceLow = RGB(255, 0, 0)
        'Set lower limit = "10"
        .ToleranceLow = 10
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA759
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis = absolute
        .TypeWarningHigh = False
        'Activate monitoring
        .CheckWarningHigh = True
        'Set barcolor = "red"
        .ColorWarningHigh = RGB(255, 0, 0)
        'Set upper limit = "75"
        .WarningHigh = 75
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA760
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis = absolute
        .TypeWarningLow = False
        'Activate monitoring
        .CheckWarningLow = True
        'Set barcolor = "magenta"
        .ColorWarningLow = RGB(255, 0, 255)
        'Set lower limit = "12"
        .WarningLow = 12
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ExampleForLanguageFonts()
'VBA761
    Dim colLangFonts As HMILanguageFonts
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    objButton.Text = "DefText"
    Set colLangFonts = objButton.LDFonts
'
    'Set font-properties for french:
    With colLangFonts.ItemByLCID(1036)
        .Family = "Courier New"
        .Bold = True
        .Italic = False
        .Underlined = True
        .Size = 12
    End With
'
    'Set font-properties for english:
    With colLangFonts.ItemByLCID(1033)
        .Family = "Times New Roman"
        .Bold = False
        .Italic = True
        .Underlined = False
        .Size = 14
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub TextListConfiguration()
'VBA762
    Dim objTextList As HMITextList
    Set objTextList = ActiveDocument.HMIObjects.AddHMIObject("myTextList", "HMITextList")
    With objTextList
        .UnselBGColor = RGB(255, 0, 0)
        .UnselTextColor = RGB(0, 0, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub TextListConfiguration()
'VBA763
    Dim objTextList As HMITextList
    Set objTextList = ActiveDocument.HMIObjects.AddHMIObject("myTextList", "HMITextList")
    With objTextList
        .UnselBGColor = RGB(255, 0, 0)
        .UnselTextColor = RGB(0, 0, 0)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub PictureWindowConfig()
'VBA764
     Dim objPicWindow As HMIPictureWindow
     Set objPicWindow = ActiveDocument.HMIObjects.AddHMIObject("PicWindow1", "HMIPictureWindow")
     With objPicWindow
          .UpdateCycle = 5
     End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA765
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .UserValue1 = 0
        .UserValue2 = 25
        .UserValue3 = 50
        .UserValue4 = 75
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA766
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .UserValue1 = 0
        .UserValue2 = 25
        .UserValue3 = 50
        .UserValue4 = 75
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA767
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .UserValue1 = 0
        .UserValue2 = 25
        .UserValue3 = 50
        .UserValue4 = 75
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GroupDisplayConfiguration()
'VBA768
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .UserValue1 = 0
        .UserValue2 = 25
        .UserValue3 = 50
        .UserValue4 = 75
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddActiveXControl()
'VBA769
    Dim objActiveXControl As HMIActiveXControl
    Set objActiveXControl = ActiveDocument.HMIObjects.AddActiveXControl("WinCC_Gauge2", "XGAUGE.XGaugeCtrl.1")
'
    'Move ActiveX-Control:
    objActiveXControl.Top = 40
    objActiveXControl.Left = 60
'
    'Modify individual properties:
    objActiveXControl.Properties("BackColor").value = RGB(255, 0, 0)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeAnalog()
'VBA770
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeAnalog
        .AnalogResultInfos.ElseCase = 200
'
        'Activate analysis of variablestate
        .VariableStateChecked = True
    End With
    With objDynDialog.VariableStateValues(1)
'
        'define a value for every state:
        .VALUE_ACCESS_FAULT = 20
        .VALUE_ADDRESS_ERROR = 30
        .VALUE_CONVERSION_ERROR = 40
        .VALUE_HANDSHAKE_ERROR = 60
        .VALUE_HARDWARE_ERROR = 70
        .VALUE_INVALID_KEY = 80
        .VALUE_MAX_LIMIT = 90
        .VALUE_MAX_RANGE = 100
        .VALUE_MIN_LIMIT = 110
        .VALUE_MIN_RANGE = 120
        .VALUE_NOT_ESTABLISHED = 130
        .VALUE_SERVERDOWN = 140
        .VALUE_STARTUP_VALUE = 150
        .VALUE_TIMEOUT = 160
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeAnalog()
'VBA771
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeAnalog
        .AnalogResultInfos.ElseCase = 200
'
        'Activate analysis of variablestate
        .VariableStateChecked = True
    End With
    With objDynDialog.VariableStateValues(1)
'
        'define a value for every state:
        .VALUE_ACCESS_FAULT = 20
        .VALUE_ADDRESS_ERROR = 30
        .VALUE_CONVERSION_ERROR = 40
        .VALUE_HANDSHAKE_ERROR = 60
        .VALUE_HARDWARE_ERROR = 70
        .VALUE_INVALID_KEY = 80
        .VALUE_MAX_LIMIT = 90
        .VALUE_MAX_RANGE = 100
        .VALUE_MIN_LIMIT = 110
        .VALUE_MIN_RANGE = 120
        .VALUE_NOT_ESTABLISHED = 130
        .VALUE_SERVERDOWN = 140
        .VALUE_STARTUP_VALUE = 150
        .VALUE_TIMEOUT = 160
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeAnalog()
'VBA772
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeAnalog
        .AnalogResultInfos.ElseCase = 200
'
        'Activate analysis of variablestate
        .VariableStateChecked = True
    End With
    With objDynDialog.VariableStateValues(1)
'
        'define a value for every state:
        .VALUE_ACCESS_FAULT = 20
        .VALUE_ADDRESS_ERROR = 30
        .VALUE_CONVERSION_ERROR = 40
        .VALUE_HANDSHAKE_ERROR = 60
        .VALUE_HARDWARE_ERROR = 70
        .VALUE_INVALID_KEY = 80
        .VALUE_MAX_LIMIT = 90
        .VALUE_MAX_RANGE = 100
        .VALUE_MIN_LIMIT = 110
        .VALUE_MIN_RANGE = 120
        .VALUE_NOT_ESTABLISHED = 130
        .VALUE_SERVERDOWN = 140
        .VALUE_STARTUP_VALUE = 150
        .VALUE_TIMEOUT = 160
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeAnalog()
'VBA773
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeAnalog
        .AnalogResultInfos.ElseCase = 200
'
        'Activate analysis of variablestate
        .VariableStateChecked = True
    End With
    With objDynDialog.VariableStateValues(1)
'
        'define a value for every state:
        .VALUE_ACCESS_FAULT = 20
        .VALUE_ADDRESS_ERROR = 30
        .VALUE_CONVERSION_ERROR = 40
        .VALUE_HANDSHAKE_ERROR = 60
        .VALUE_HARDWARE_ERROR = 70
        .VALUE_INVALID_KEY = 80
        .VALUE_MAX_LIMIT = 90
        .VALUE_MAX_RANGE = 100
        .VALUE_MIN_LIMIT = 110
        .VALUE_MIN_RANGE = 120
        .VALUE_NOT_ESTABLISHED = 130
        .VALUE_SERVERDOWN = 140
        .VALUE_STARTUP_VALUE = 150
        .VALUE_TIMEOUT = 160
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeAnalog()
'VBA774
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeAnalog
        .AnalogResultInfos.ElseCase = 200
'
        'Activate analysis of variablestate
        .VariableStateChecked = True
    End With
    With objDynDialog.VariableStateValues(1)
'
        'define a value for every state:
        .VALUE_ACCESS_FAULT = 20
        .VALUE_ADDRESS_ERROR = 30
        .VALUE_CONVERSION_ERROR = 40
        .VALUE_HANDSHAKE_ERROR = 60
        .VALUE_HARDWARE_ERROR = 70
        .VALUE_INVALID_KEY = 80
        .VALUE_MAX_LIMIT = 90
        .VALUE_MAX_RANGE = 100
        .VALUE_MIN_LIMIT = 110
        .VALUE_MIN_RANGE = 120
        .VALUE_NOT_ESTABLISHED = 130
        .VALUE_SERVERDOWN = 140
        .VALUE_STARTUP_VALUE = 150
        .VALUE_TIMEOUT = 160
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeAnalog()
'VBA775
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeAnalog
        .AnalogResultInfos.ElseCase = 200
'
        'Activate analysis of variablestate
        .VariableStateChecked = True
    End With
    With objDynDialog.VariableStateValues(1)
'
        'define a value for every state:
        .VALUE_ACCESS_FAULT = 20
        .VALUE_ADDRESS_ERROR = 30
        .VALUE_CONVERSION_ERROR = 40
        .VALUE_HANDSHAKE_ERROR = 60
        .VALUE_HARDWARE_ERROR = 70
        .VALUE_INVALID_KEY = 80
        .VALUE_MAX_LIMIT = 90
        .VALUE_MAX_RANGE = 100
        .VALUE_MIN_LIMIT = 110
        .VALUE_MIN_RANGE = 120
        .VALUE_NOT_ESTABLISHED = 130
        .VALUE_SERVERDOWN = 140
        .VALUE_STARTUP_VALUE = 150
        .VALUE_TIMEOUT = 160
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeAnalog()
'VBA776
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeAnalog
        .AnalogResultInfos.ElseCase = 200
'
        'Activate analysis of variablestate
        .VariableStateChecked = True
    End With
    With objDynDialog.VariableStateValues(1)
'
        'define a value for every state:
        .VALUE_ACCESS_FAULT = 20
        .VALUE_ADDRESS_ERROR = 30
        .VALUE_CONVERSION_ERROR = 40
        .VALUE_HANDSHAKE_ERROR = 60
        .VALUE_HARDWARE_ERROR = 70
        .VALUE_INVALID_KEY = 80
        .VALUE_MAX_LIMIT = 90
        .VALUE_MAX_RANGE = 100
        .VALUE_MIN_LIMIT = 110
        .VALUE_MIN_RANGE = 120
        .VALUE_NOT_ESTABLISHED = 130
        .VALUE_SERVERDOWN = 140
        .VALUE_STARTUP_VALUE = 150
        .VALUE_TIMEOUT = 160
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeAnalog()
'VBA777
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeAnalog
        .AnalogResultInfos.ElseCase = 200
'
        'Activate analysis of variablestate
        .VariableStateChecked = True
    End With
    With objDynDialog.VariableStateValues(1)
'
        'define a value for every state:
        .VALUE_ACCESS_FAULT = 20
        .VALUE_ADDRESS_ERROR = 30
        .VALUE_CONVERSION_ERROR = 40
        .VALUE_HANDSHAKE_ERROR = 60
        .VALUE_HARDWARE_ERROR = 70
        .VALUE_INVALID_KEY = 80
        .VALUE_MAX_LIMIT = 90
        .VALUE_MAX_RANGE = 100
        .VALUE_MIN_LIMIT = 110
        .VALUE_MIN_RANGE = 120
        .VALUE_NOT_ESTABLISHED = 130
        .VALUE_SERVERDOWN = 140
        .VALUE_STARTUP_VALUE = 150
        .VALUE_TIMEOUT = 160
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeAnalog()
'VBA778
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeAnalog
        .AnalogResultInfos.ElseCase = 200
'
        'Activate analysis of variablestate
        .VariableStateChecked = True
    End With
    With objDynDialog.VariableStateValues(1)
'
        'define a value for every state:
        .VALUE_ACCESS_FAULT = 20
        .VALUE_ADDRESS_ERROR = 30
        .VALUE_CONVERSION_ERROR = 40
        .VALUE_HANDSHAKE_ERROR = 60
        .VALUE_HARDWARE_ERROR = 70
        .VALUE_INVALID_KEY = 80
        .VALUE_MAX_LIMIT = 90
        .VALUE_MAX_RANGE = 100
        .VALUE_MIN_LIMIT = 110
        .VALUE_MIN_RANGE = 120
        .VALUE_NOT_ESTABLISHED = 130
        .VALUE_SERVERDOWN = 140
        .VALUE_STARTUP_VALUE = 150
        .VALUE_TIMEOUT = 160
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeAnalog()
'VBA779
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeAnalog
        .AnalogResultInfos.ElseCase = 200
'
        'Activate analysis of variablestate
        .VariableStateChecked = True
    End With
    With objDynDialog.VariableStateValues(1)
'
        'define a value for every state:
        .VALUE_ACCESS_FAULT = 20
        .VALUE_ADDRESS_ERROR = 30
        .VALUE_CONVERSION_ERROR = 40
        .VALUE_HANDSHAKE_ERROR = 60
        .VALUE_HARDWARE_ERROR = 70
        .VALUE_INVALID_KEY = 80
        .VALUE_MAX_LIMIT = 90
        .VALUE_MAX_RANGE = 100
        .VALUE_MIN_LIMIT = 110
        .VALUE_MIN_RANGE = 120
        .VALUE_NOT_ESTABLISHED = 130
        .VALUE_SERVERDOWN = 140
        .VALUE_STARTUP_VALUE = 150
        .VALUE_TIMEOUT = 160
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeAnalog()
'VBA780
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeAnalog
        .AnalogResultInfos.ElseCase = 200
'
        'Activate analysis of variablestate
        .VariableStateChecked = True
    End With
    With objDynDialog.VariableStateValues(1)
'
        'define a value for every state:
        .VALUE_ACCESS_FAULT = 20
        .VALUE_ADDRESS_ERROR = 30
        .VALUE_CONVERSION_ERROR = 40
        .VALUE_HANDSHAKE_ERROR = 60
        .VALUE_HARDWARE_ERROR = 70
        .VALUE_INVALID_KEY = 80
        .VALUE_MAX_LIMIT = 90
        .VALUE_MAX_RANGE = 100
        .VALUE_MIN_LIMIT = 110
        .VALUE_MIN_RANGE = 120
        .VALUE_NOT_ESTABLISHED = 130
        .VALUE_SERVERDOWN = 140
        .VALUE_STARTUP_VALUE = 150
        .VALUE_TIMEOUT = 160
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeAnalog()
'VBA781
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeAnalog
        .AnalogResultInfos.ElseCase = 200
'
        'Activate analysis of variablestate
        .VariableStateChecked = True
    End With
    With objDynDialog.VariableStateValues(1)
'
        'define a value for every state:
        .VALUE_ACCESS_FAULT = 20
        .VALUE_ADDRESS_ERROR = 30
        .VALUE_CONVERSION_ERROR = 40
        .VALUE_HANDSHAKE_ERROR = 60
        .VALUE_HARDWARE_ERROR = 70
        .VALUE_INVALID_KEY = 80
        .VALUE_MAX_LIMIT = 90
        .VALUE_MAX_RANGE = 100
        .VALUE_MIN_LIMIT = 110
        .VALUE_MIN_RANGE = 120
        .VALUE_NOT_ESTABLISHED = 130
        .VALUE_SERVERDOWN = 140
        .VALUE_STARTUP_VALUE = 150
        .VALUE_TIMEOUT = 160
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeAnalog()
'VBA782
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeAnalog
        .AnalogResultInfos.ElseCase = 200
'
        'Activate analysis of variablestate
        .VariableStateChecked = True
    End With
    With objDynDialog.VariableStateValues(1)
'
        'define a value for every state:
        .VALUE_ACCESS_FAULT = 20
        .VALUE_ADDRESS_ERROR = 30
        .VALUE_CONVERSION_ERROR = 40
        .VALUE_HANDSHAKE_ERROR = 60
        .VALUE_HARDWARE_ERROR = 70
        .VALUE_INVALID_KEY = 80
        .VALUE_MAX_LIMIT = 90
        .VALUE_MAX_RANGE = 100
        .VALUE_MIN_LIMIT = 110
        .VALUE_MIN_RANGE = 120
        .VALUE_NOT_ESTABLISHED = 130
        .VALUE_SERVERDOWN = 140
        .VALUE_STARTUP_VALUE = 150
        .VALUE_TIMEOUT = 160
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeAnalog()
'VBA783
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeAnalog
        .AnalogResultInfos.ElseCase = 200
'
        'Activate analysis of variablestate
        .VariableStateChecked = True
    End With
    With objDynDialog.VariableStateValues(1)
'
        'define a value for every state:
        .VALUE_ACCESS_FAULT = 20
        .VALUE_ADDRESS_ERROR = 30
        .VALUE_CONVERSION_ERROR = 40
        .VALUE_HANDSHAKE_ERROR = 60
        .VALUE_HARDWARE_ERROR = 70
        .VALUE_INVALID_KEY = 80
        .VALUE_MAX_LIMIT = 90
        .VALUE_MAX_RANGE = 100
        .VALUE_MIN_LIMIT = 110
        .VALUE_MIN_RANGE = 120
        .VALUE_NOT_ESTABLISHED = 130
        .VALUE_SERVERDOWN = 140
        .VALUE_STARTUP_VALUE = 150
        .VALUE_TIMEOUT = 160
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeAnalog()
'VBA785
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeAnalog
        .AnalogResultInfos.ElseCase = 200
'
        'Activate analysis of variablestate
        .VariableStateChecked = True
    End With
    With objDynDialog.VariableStateValues(1)
'
        'define a value for every state:
        .VALUE_ACCESS_FAULT = 20
        .VALUE_ADDRESS_ERROR = 30
        .VALUE_CONVERSION_ERROR = 40
        .VALUE_HANDSHAKE_ERROR = 60
        .VALUE_HARDWARE_ERROR = 70
        .VALUE_INVALID_KEY = 80
        .VALUE_MAX_LIMIT = 90
        .VALUE_MAX_RANGE = 100
        .VALUE_MIN_LIMIT = 110
        .VALUE_MIN_RANGE = 120
        .VALUE_NOT_ESTABLISHED = 130
        .VALUE_SERVERDOWN = 140
        .VALUE_STARTUP_VALUE = 150
        .VALUE_TIMEOUT = 160
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddDynamicDialogToCircleRadiusTypeAnalog()
'VBA786
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeAnalog
        .AnalogResultInfos.ElseCase = 200
'
        'Activate analysis of variablestate
        .VariableStateChecked = True
    End With
    With objDynDialog.VariableStateValues(1)
'
        'define a value for every state:
        .VALUE_ACCESS_FAULT = 20
        .VALUE_ADDRESS_ERROR = 30
        .VALUE_CONVERSION_ERROR = 40
        .VALUE_HANDSHAKE_ERROR = 60
        .VALUE_HARDWARE_ERROR = 70
        .VALUE_INVALID_KEY = 80
        .VALUE_MAX_LIMIT = 90
        .VALUE_MAX_RANGE = 100
        .VALUE_MIN_LIMIT = 110
        .VALUE_MIN_RANGE = 120
        .VALUE_NOT_ESTABLISHED = 130
        .VALUE_SERVERDOWN = 140
        .VALUE_STARTUP_VALUE = 150
        .VALUE_TIMEOUT = 160
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DynamicWithVariableTrigger()
'VBA787
    Dim objVBScript As HMIScriptInfo
    Dim objVarTrigger As HMIVariableTrigger
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_VariableTrigger", "HMICircle")
    Set objVBScript = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeVBScript)
    With objVBScript
        'Triggername and cycletime are defined by add-methode
        Set objVarTrigger = .Trigger.VariableTriggers.Add("VarTrigger", hmiVariableCycleType_10s)
        .SourceCode = ""
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GetVarName()
'VBA788
    Dim objVBScript As HMIScriptInfo
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.Item("Circle_VariableTrigger")
    Set objVBScript = objCircle.Radius.Dynamic
    With objVBScript
        'Reading out of variablename
        MsgBox "The radius is dynamicabled with: " & .Trigger.VariableTriggers.Item(1).VarName
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowVBAVersion()
'VBA789
    MsgBox Application.VBAVersion
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowVersionOfGraphicsDesigner()
'VBA791
    MsgBox Application.Version
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddView()
'VBA792
    Dim objView As HMIView
    Set objView = ActiveDocument.Views.Add
    objView.Activate
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HideCircleInRuntime()
'VBA793
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("myCircle", "HMICircle")
    objCircle.Visible = False
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA794
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis = absolute
        .TypeWarningHigh = False
        'Activate monitoring
        .CheckWarningHigh = True
        'Set barcolor = "red"
        .ColorWarningHigh = RGB(255, 0, 0)
        'Set upper limit = "75"
        .WarningHigh = 75
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphLimitConfiguration()
'VBA795
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        'Set analysis = absolute
        .TypeWarningLow = False
        'Activate monitoring
        .CheckWarningLow = True
        'Set barcolor = "magenta"
        .ColorWarningLow = RGB(255, 0, 255)
        'Set lower limit = "12"
        .WarningLow = 75
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ApplySameWidthToSelectedObjects()
'VBA796
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim objEllipse As HMIEllipse
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    Set objEllipse = ActiveDocument.HMIObjects.AddHMIObject("sEllipse", "HMIEllipse")
    With objCircle
        .Top = 30
        .Left = 0
        .Width = 15
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 42
        .Width = 40
        .Selected = True
    End With
    With objEllipse
        .Top = 48
        .Left = 162
        .Width = 120
        .BackColor = RGB(255, 0, 0)
        .Selected = True
    End With
    MsgBox "Objects selected!"
    ActiveDocument.Selection.SameWidth
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ApplicationWindowConfig()
'VBA797
    Dim objAppWindow As HMIApplicationWindow
    Set objAppWindow = ActiveDocument.HMIObjects.AddHMIObject("AppWindow", "HMIApplicationWindow")
    With objAppWindow
        .Caption = True
        .CloseButton = False
        .Height = 200
        .Left = 10
        .MaximizeButton = True
        .Moveable = False
        .OnTop = True
        .Sizeable = True
        .Top = 20
        .Visible = True
        .Width = 250
        .WindowBorder = True
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowWindowState()
'VBA798
    Dim strState As String
    Select Case Application.WindowState
        Case 0
            strState = "The application-window is maximized"
        Case 1
            strState = "The applicationwindow is minimized"
        Case 2
            strState = "The application-window has a userdefined size"
    End Select
    MsgBox strState
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphConfiguration()
'VBA799
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .Scaling = True
        .ScalingType = 2
        .ZeroPoint = 50
        .ZeroPointValue = 0
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BarGraphConfiguration()
'VBA800
    Dim objBarGraph As HMIBarGraph
    Set objBarGraph = ActiveDocument.HMIObjects.AddHMIObject("Bar1", "HMIBarGraph")
    With objBarGraph
        .Scaling = True
        .ScalingType = 2
        .ZeroPointValue = 0
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateViewFromActiveDocument()
'VBA801
    Dim objView As HMIView
    Set objView = ActiveDocument.Views.Add
    objView.Zoom = 50
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Document_Opened(CancelForwarding As Boolean)
'VBA802
    MsgBox ActiveDocument.Hide
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowLayerWithNumbers()
'VBA803
    Dim colLayers As HMILayers
    Dim objLayer As HMILayer
    Dim iAnswer As Integer
    Dim iIndex As Integer
    iIndex = 1
    Set colLayers = ActiveDocument.Layers
    For Each objLayer In colLayers
        iAnswer = MsgBox("Layername: " & objLayer & vbCrLf & "Layernumber: " & objLayer.Number & vbCrLf & "Layersindex: " & iIndex, vbOKCancel)
        iIndex = iIndex + 1
        If vbCancel = iAnswer Then Exit For
    Next objLayer
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddOLEObjectDirect()
'VBA804
    Dim objOLEObject As HMIOLEObject
    Set objOLEObject = ActiveDocument.HMIObjects.AddOLEObject("OLEObject1", "Wordpad.Document.1", hmiOLEObjectCreationTypeDirect, True)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddOLEObjectByLink()
'VBA805
    Dim objOLEObject As HMIOLEObject
    Dim strFilename As String
'
    'Add OLEObject by filename. In this case, the filename has to
    'contain filename and path.
    'Replace the definition of strFilename with a filename with path
    'existing on your system
    strFilename = Application.ApplicationDataPath & "Test.bmp"
    Set objOLEObject = ActiveDocument.HMIObjects.AddOLEObject("OLEObject1", strFilename, hmiOLEObjectCreationTypeByLink, False)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddOLEObjectByLinkWithReference()
'VBA805
    Dim objOLEObject As HMIOLEObject
    Dim strFilename As String
'
    'Add OLEObject by filename. In this case, the filename has to
    'contain filename and path.
    'Replace the definition of strFilename with a filename with path
    'existing on your system
    strFilename = Application.ApplicationDataPath & "Test.bmp"
    Set objOLEObject = ActiveDocument.HMIObjects.AddOLEObject("OLEObject1", strFilename, hmiOLEObjectCreationTypeByLinkWithReference, True)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
