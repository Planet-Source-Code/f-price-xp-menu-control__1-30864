VERSION 5.00
Begin VB.UserControl XPUIMenuControl 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "XPUIMenuControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim floyd As New clsMenu
'Property Variables:
Dim m_separatorcolor As OLE_COLOR
Dim m_MenuBorderColor As OLE_COLOR
Dim m_MenuBackColor As OLE_COLOR
Dim m_MenuImageBackColor As OLE_COLOR
Dim m_MenuItemHotColor As OLE_COLOR
Dim m_MenuItemHotBorderColor As OLE_COLOR
Dim m_ImageList As Object
Dim m_Menus As New DinkITXPUIMenus.XPUIMenus
Private WithEvents ActiveMenu As XPUIMenu
Attribute ActiveMenu.VB_VarHelpID = -1
'Default Property Values:
Const m_def_SeparatorColor = &HB8C2C5
Const m_def_MenuBorderColor = &H808080
Const m_def_MenuBackColor = vbWhite
Const m_def_MenuImageBackColor = &HDEEDEF
Const m_def_MenuItemHotColor = &HD2BDB6
Const m_def_MenuItemHotBorderColor = &H6A240A

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Extender
Public Property Get Extender() As Object
Attribute Extender.VB_Description = "Returns the Extender object for this control which allows access to the properties of the control that are kept track of by the container."
    Set Extender = UserControl.Extender
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get HWND() As Long
Attribute HWND.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    HWND = UserControl.HWND
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=9,0,0,0
Public Property Get Menus() As DinkITXPUIMenus.XPUIMenus
    Set Menus = m_Menus
End Property

Public Property Set Menus(ByVal New_Menus As DinkITXPUIMenus.XPUIMenus)
    Set m_Menus = New_Menus
    PropertyChanged "Menus"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=9,0,0,0
Public Property Get ImageList() As Object
    Set ImageList = m_ImageList
End Property

Public Property Set ImageList(ByVal New_ImageList As Object)
    Set m_ImageList = New_ImageList
    PropertyChanged "ImageList"
End Property

Private Sub ActiveMenu_Click(Menu As XPUIMenu, MenuItem As MenuItem)
    MsgBox "Clicked : " & MenuItem.Caption
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set m_ImageList = PropBag.ReadProperty("ImageList", Nothing)
    
    m_MenuBorderColor = PropBag.ReadProperty("MenuBorderColor", m_def_MenuBorderColor)
    m_MenuBackColor = PropBag.ReadProperty("MenuBackColor", m_def_MenuBackColor)
    m_MenuImageBackColor = PropBag.ReadProperty("MenuImageBackColor", m_def_MenuImageBackColor)
    m_MenuItemHotColor = PropBag.ReadProperty("MenuItemHotColor", m_def_MenuItemHotColor)
    m_MenuItemHotBorderColor = PropBag.ReadProperty("MenuItemHotBorderColor", m_def_MenuItemHotBorderColor)
    m_separatorcolor = PropBag.ReadProperty("SeparatorColor", m_def_SeparatorColor)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("ImageList", m_ImageList, Nothing)
    Call PropBag.WriteProperty("MenuBorderColor", m_MenuBorderColor, m_def_MenuBorderColor)
    Call PropBag.WriteProperty("MenuBackColor", m_MenuBackColor, m_def_MenuBackColor)
    Call PropBag.WriteProperty("MenuImageBackColor", m_MenuImageBackColor, m_def_MenuImageBackColor)
    Call PropBag.WriteProperty("MenuItemHotColor", m_MenuItemHotColor, m_def_MenuItemHotColor)
    Call PropBag.WriteProperty("MenuItemHotBorderColor", m_MenuItemHotBorderColor, m_def_MenuItemHotBorderColor)
    Call PropBag.WriteProperty("SeparatorColor", m_separatorcolor, m_def_SeparatorColor)
End Sub

Public Sub ShowMenu(Menu As String, x As String, y As String)
    Set ActiveMenu = Menus(Menu)
    Set ActiveMenu.ImageList = m_ImageList
    ActiveMenu.MenuBorderColor = m_MenuBorderColor
    ActiveMenu.MenuItemHotColor = m_MenuItemHotColor
    ActiveMenu.separatorcolor = m_separatorcolor
    ActiveMenu.MenuBackColor = m_MenuBackColor
    ActiveMenu.MenuImageBackColor = m_MenuImageBackColor
    ActiveMenu.MenuItemBorderColor = m_MenuItemHotBorderColor
    ActiveMenu.ShowMenu CLng(x), CLng(y)
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MenuBorderColor() As OLE_COLOR
    MenuBorderColor = m_MenuBorderColor
End Property

Public Property Let MenuBorderColor(ByVal New_MenuBorderColor As OLE_COLOR)
    m_MenuBorderColor = New_MenuBorderColor
    PropertyChanged "MenuBorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MenuBackColor() As OLE_COLOR
    MenuBackColor = m_MenuBackColor
End Property

Public Property Let MenuBackColor(ByVal New_MenuBackColor As OLE_COLOR)
    m_MenuBackColor = New_MenuBackColor
    PropertyChanged "MenuBackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MenuImageBackColor() As OLE_COLOR
    MenuImageBackColor = m_MenuImageBackColor
End Property

Public Property Let MenuImageBackColor(ByVal New_MenuImageBackColor As OLE_COLOR)
    m_MenuImageBackColor = New_MenuImageBackColor
    PropertyChanged "MenuImageBackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MenuItemHotColor() As OLE_COLOR
    MenuItemHotColor = m_MenuItemHotColor
End Property

Public Property Let MenuItemHotColor(ByVal New_MenuItemHotColor As OLE_COLOR)
    m_MenuItemHotColor = New_MenuItemHotColor
    PropertyChanged "MenuItemHotColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MenuItemHotBorderColor() As OLE_COLOR
    MenuItemHotBorderColor = m_MenuItemHotBorderColor
End Property

Public Property Let MenuItemHotBorderColor(ByVal New_MenuItemHotBorderColor As OLE_COLOR)
    m_MenuItemHotBorderColor = New_MenuItemHotBorderColor
    PropertyChanged "MenuItemHotBorderColor"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_MenuBorderColor = m_def_MenuBorderColor
    m_MenuBackColor = m_def_MenuBackColor
    m_MenuImageBackColor = m_def_MenuImageBackColor
    m_MenuItemHotColor = m_def_MenuItemHotColor
    m_MenuItemHotBorderColor = m_def_MenuItemHotBorderColor
    m_separatorcolor = m_def_SeparatorColor
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get separatorcolor() As OLE_COLOR
    separatorcolor = m_separatorcolor
End Property

Public Property Let separatorcolor(ByVal New_SeparatorColor As OLE_COLOR)
    m_separatorcolor = New_SeparatorColor
    PropertyChanged "SeparatorColor"
End Property

