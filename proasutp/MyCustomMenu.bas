'       (с) 2016 http://www.proasutp.com
'
'       Пример создания пользовательского меню в графическом дизайнере WinCC (WinCC Graphics Designer).
'       В примере показан вызов пользовательской функции - подсчета количества графических объектов на графической странице.
'
'      Использование: Данный код можно вставить для демонстрации работы, например, на уровне текущей открытой графической страницы
'                     (Tools->Macros->Visual Basic Editor ->VBAProject(X)->Graphics Designer Objects->ThisDocument (PageName.Pdl))
'
'       Вход:
'                Нет
'        Выход:
'                Нет

Dim WithEvents theApp As grafexe.Application

Private Const m_sMenuTitle = "Автоматизация задач"
Private Const m_sItemMenuTitle = "Подсчет количества графических объектов на странице"

' Событие, возникающие при открытии графической страницы в Graphics Designer

Private Sub Document_Opened(CancelForwarding As Boolean)
    Set theApp = grafexe.Application
    CreateMyCustomMenu
End Sub

' Создание пользовательского пункта меню в главном меню Graphics Designer

Private Sub CreateMyCustomMenu()
   
    Dim objMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem

    Set objMenu = Application.CustomMenus.InsertMenu(1, m_sMenuTitle, m_sMenuTitle)
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(1, m_sItemMenuTitle, m_sItemMenuTitle)

End Sub

' Отработка события выбора элемента меню

Private Sub theApp_MenuItemClicked(ByVal MenuItem As IHMIMenuItem)

    Dim objMenuClicked As HMIMenuItem
    Dim c As Long
       
    Set objMenuClicked = MenuItem
   
    Select Case objMenuClicked.Key
         Case m_sItemMenuTitle
            c = GetGraphicsObjectCount ()
            MsgBox "Количество графических объектов, шт: " & c, vbOKOnly, m_sMenuTitle
    End Select

End Sub

' Получить количество графических объектов на текущей графической странице

Private Function GetGraphicsObjectCount () As Long
    GetGraphicsObjectCount = theApp.ActiveDocument.HMIObjects.Count
End Function