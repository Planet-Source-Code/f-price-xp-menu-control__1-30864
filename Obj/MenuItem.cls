VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"XPUIMenu"
Attribute VB_Ext_KEY = "Member1" ,"XPUIMenu"
Option Explicit
Private Const DST_BITMAP = &H4
Private Const DST_ICON = &H3
Private Const DST_COMPLEX = 16
Private Const DSS_DISABLED = &H20
Private Const DSS_MONO = &H80
Private Declare Function DrawState Lib "User32" Alias "DrawStateA" (ByVal HDC As Long, ByVal hBr As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As Long
Private Const DSS_NORMAL = &H0

Private mvarIcon As Long
Private mvarCaption As String
Private mvarRadioOption As Boolean
Private mvarCheckOption As Boolean
Private mvarOptionSelected As Boolean
Private mvarDisabled As Boolean
Private m_Key As String
Private mvarXPUIMenu As XPUIMenu
'local variable(s) to hold property value(s)
Private mvarSeperator As Boolean 'local copy
Private m_Item As Integer
Private m_MouseOver As Boolean
Private m_Top As Long
Private m_Left As Long
Private m_Right As Long
Private m_Bottom As Long
Private m_MenuObj As Object
Private m_ImageList As Object
Private m_Popped As Boolean
Private m_ExpandImage As StdPicture
Private m_MenuBackColor As OLE_COLOR
Private m_MenuImageBackColor As OLE_COLOR
Private m_MenuItemHotColor As OLE_COLOR
Private m_MenuItemBorderColor As OLE_COLOR
Private m_separatorcolor As OLE_COLOR

Friend Property Get separatorcolor() As OLE_COLOR
    separatorcolor = m_separatorcolor
End Property

Friend Property Let separatorcolor(ByVal Value As OLE_COLOR)
    m_separatorcolor = Value
End Property

Friend Property Get MenuItemBorderColor() As OLE_COLOR
    MenuItemBorderColor = m_MenuItemBorderColor
End Property

Friend Property Let MenuItemBorderColor(ByVal Value As OLE_COLOR)
    m_MenuItemBorderColor = Value
End Property

Friend Property Get MenuItemHotColor() As OLE_COLOR
    MenuItemHotColor = m_MenuItemHotColor
End Property

Friend Property Let MenuItemHotColor(ByVal Value As OLE_COLOR)
    m_MenuItemHotColor = Value
End Property

Friend Property Get MenuImageBackColor() As OLE_COLOR
    MenuImageBackColor = m_MenuImageBackColor
End Property

Friend Property Let MenuImageBackColor(ByVal Value As OLE_COLOR)
    m_MenuImageBackColor = Value
End Property

Friend Property Get MenuBackColor() As OLE_COLOR
    MenuBackColor = m_MenuBackColor
End Property

Friend Property Let MenuBackColor(ByVal Value As OLE_COLOR)
    m_MenuBackColor = Value
End Property

Friend Property Get ExpandImage() As StdPicture
    Set ExpandImage = m_ExpandImage
End Property

Friend Property Set ExpandImage(ByVal Value As StdPicture)
    Set m_ExpandImage = Value
End Property

Friend Property Get Popped() As Boolean
    Popped = m_Popped
End Property

Friend Property Let Popped(ByVal Value As Boolean)
    m_Popped = Value
End Property

Friend Property Get ImageList() As Object
    Set ImageList = m_ImageList
End Property

Friend Property Set ImageList(ByVal Value As Object)
    Set m_ImageList = Value
End Property


Friend Property Get MenuObj() As Object
    Set MenuObj = m_MenuObj
End Property

Friend Property Set MenuObj(ByVal Value As Object)
    Set m_MenuObj = Value
End Property

Friend Property Get Bottom() As Long
    Bottom = m_Bottom
End Property

Friend Property Let Bottom(ByVal Value As Long)
    m_Bottom = Value
End Property

Friend Property Get Right() As Long
    Right = m_Right
End Property

Friend Property Let Right(ByVal Value As Long)
    m_Right = Value
End Property

Friend Property Get Left() As Long
    Left = m_Left
End Property

Friend Property Let Left(ByVal Value As Long)
    m_Left = Value
End Property

Friend Property Get Top() As Long
    Top = m_Top
End Property

Friend Property Let Top(ByVal Value As Long)
    m_Top = Value
End Property

Friend Property Get MouseOver() As Boolean
    MouseOver = m_MouseOver
End Property

Friend Property Let MouseOver(ByVal Value As Boolean)
    m_MouseOver = Value
    pDrawItem
'    Dim oPT As POINTAPI
'    ClientToScreen MenuObj.HWND, oPT
'
'    If Me.XPUIMenu.MenuItems.Count > 0 Then
'
'        If m_MouseOver Then
'            Set oNewMenu = Me.XPUIMenu
'            oNewMenu.ShowMenu oPT.X + ((m_Right - m_Left) + 3), oPT.Y + m_Top
'            oNewMenu.MouseOver = True
'        Else
'            If oNewMenu.MouseOver Then
'
'            Else
'                oNewMenu.Term
'                Set oNewMenu = Nothing
'                Debug.Print "Nothing"
'            End If
'        End If
'    Else
'        If oNewMenu Is Nothing Then
'
'        Else
'            oNewMenu.Term
'            Set oNewMenu = Nothing
'        End If
'    End If
End Property

Public Property Get Item() As Integer
    Item = m_Item
End Property

Public Property Let Item(ByVal Value As Integer)
    m_Item = Value
End Property
Public Property Let Seperator(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Seperator = 5
    mvarSeperator = vData
End Property


Public Property Get Seperator() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Seperator
    Seperator = mvarSeperator
End Property

Public Property Get XPUIMenu() As XPUIMenu
    If mvarXPUIMenu Is Nothing Then
        Set mvarXPUIMenu = New XPUIMenu
    End If


    Set XPUIMenu = mvarXPUIMenu
End Property


Public Property Set XPUIMenu(vData As XPUIMenu)
    Set mvarXPUIMenu = vData
End Property

Private Sub Class_Terminate()
  Set mvarXPUIMenu = Nothing
    
End Sub

Public Property Get Key() As String
    Key = m_Key
End Property

Public Property Let Key(ByVal Value As String)
    m_Key = Value
End Property

Public Property Let Disabled(ByVal vData As Boolean)
    mvarDisabled = vData
End Property

Public Property Get Disabled() As Boolean
    Disabled = mvarDisabled
End Property

Public Property Let OptionSelected(ByVal vData As Boolean)
    mvarOptionSelected = vData
End Property

Public Property Get OptionSelected() As Boolean
    OptionSelected = mvarOptionSelected
End Property

Public Property Let CheckOption(ByVal vData As Boolean)
    mvarCheckOption = vData
End Property

Public Property Get CheckOption() As Boolean
    CheckOption = mvarCheckOption
End Property

Public Property Let Caption(ByVal vData As String)
    mvarCaption = vData
End Property

Public Property Get Caption() As String
    Caption = mvarCaption
End Property

Public Property Let Icon(ByVal vData As Long)
    mvarIcon = vData
End Property

Public Property Get Icon() As Long
    Icon = mvarIcon
End Property

Friend Sub pDrawItem()
    On Error Resume Next
    Dim oRCBounds As Rect
    oRCBounds.Left = m_Left
    oRCBounds.Right = m_Right
    oRCBounds.Top = m_Top
    oRCBounds.Bottom = m_Bottom
    
    If mvarSeperator Then
        
        oRCBounds.Bottom = oRCBounds.Top + 1
        oRCBounds.Left = 28
        BoxRect3DDCex MenuObj.HDC, oRCBounds, m_separatorcolor, m_separatorcolor, m_separatorcolor ' &HB8C2C5, &HB8C2C5, &HB8C2C5
    Else
        If m_Popped = True Or (m_MouseOver And mvarDisabled = False) Then
            BoxRect3DDCex MenuObj.HDC, oRCBounds, m_MenuItemBorderColor, m_MenuItemBorderColor, m_MenuItemHotColor
            If Icon = 0 Then
        
            Else
                If Icon > ImageList.listimages.Count Then
                
                Else
                    
                    fDrawPicture ImageList.listimages(Icon).Picture, oRCBounds.Left + 4, (oRCBounds.Bottom - ((oRCBounds.Bottom - oRCBounds.Top) / 2) - 8), 16, 16, True, mvarDisabled = True
                    If Not mvarDisabled Then
                        fDrawPicture ImageList.listimages(Icon).Picture, oRCBounds.Left + 2, (oRCBounds.Bottom - ((oRCBounds.Bottom - oRCBounds.Top) / 2) - 10), 16, 16, False, mvarDisabled = True
                    End If
                End If
            End If
            If Me.XPUIMenu.MenuItems.Count > 0 Then
                fDrawPicture m_ExpandImage, oRCBounds.Right - 16, (oRCBounds.Bottom - ((oRCBounds.Bottom - oRCBounds.Top) / 2) - 8), 16, 16, False, False
            End If
        Else
            Dim iRVl As Integer
            Dim iLVl As Integer
            iLVl = oRCBounds.Left
            iRVl = oRCBounds.Right
            oRCBounds.Right = oRCBounds.Left + 23
            BoxRect3DDCex MenuObj.HDC, oRCBounds, m_MenuImageBackColor, m_MenuImageBackColor, m_MenuImageBackColor  '&HDEEDEF, &HDEEDEF, &HDEEDEF
            oRCBounds.Left = oRCBounds.Right
            oRCBounds.Right = iRVl
            BoxRect3DDCex MenuObj.HDC, oRCBounds, m_MenuBackColor, m_MenuBackColor, m_MenuBackColor
            oRCBounds.Left = iLVl
            If Icon = 0 Then
        
            Else
                If Icon > ImageList.listimages.Count Then
                
                Else
                    fDrawPicture ImageList.listimages(Icon).Picture, oRCBounds.Left + 4, (oRCBounds.Bottom - ((oRCBounds.Bottom - oRCBounds.Top) / 2) - 8), 16, 16, False, mvarDisabled = True
                    'fDrawPicture ImageList.listimages(Icon).Picture, oRCBounds.Left + 2, oRCBounds.Top + 2, 16, 16, False, mvarDisabled = True
                End If
            End If
           If Me.XPUIMenu.MenuItems.Count > 0 Then
                fDrawPicture m_ExpandImage, oRCBounds.Right - 16, (oRCBounds.Bottom - ((oRCBounds.Bottom - oRCBounds.Top) / 2) - 8), 16, 16, False, False
            End If
        End If
        oRCBounds.Left = 29
        If mvarDisabled Then
            Dim oRigForeColor As OLE_COLOR
            oRigForeColor = MenuObj.ForeColor
            MenuObj.ForeColor = &HB8C2C5
            
            DrawText MenuObj.HDC, mvarCaption, Len(mvarCaption), oRCBounds, 564
            
            MenuObj.ForeColor = oRigForeColor
        Else
            DrawText MenuObj.HDC, mvarCaption, Len(mvarCaption), oRCBounds, 564
        End If
    End If
    
End Sub


Private Sub fDrawPicture( _
    ByRef m_Picture As StdPicture, _
    ByVal x As Long, _
    ByVal y As Long, ByVal W As Long, ByVal H As Long, _
    ByVal bShadow As Boolean, Optional Disabled As Boolean = False)
On Error Resume Next
     
    Dim lFlags As Long
    Dim hBrush As Long
         
    Select Case m_Picture.Type
        Case vbPicTypeBitmap
            lFlags = DST_BITMAP
        Case vbPicTypeIcon
            lFlags = DST_ICON
        Case Else
            lFlags = DST_COMPLEX
    End Select

    If bShadow Then
        hBrush = CreateSolidBrush(&H9C8181)  'RGB(128, 128, 128))
    End If
    If Disabled Then
     DrawState MenuObj.HDC, IIf(bShadow, hBrush, 0), 0, m_Picture.Handle, 0, x, y, W, _
        H, _
        lFlags Or DSS_DISABLED
    Else
     DrawState MenuObj.HDC, IIf(bShadow, hBrush, 0), 0, m_Picture.Handle, 0, x, y, W, _
        H, _
        lFlags Or IIf(bShadow, DSS_MONO, DSS_NORMAL)
     End If
    If bShadow Then
        DeleteObject hBrush
    End If
     
End Sub


