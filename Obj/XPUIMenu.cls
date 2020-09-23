VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XPUIMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MenuItems"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
Private Declare Function GetFocus Lib "User32" () As Long

Private mvarMenuItems As MenuItems
Private m_Key As String
Dim oActivePopup As XPUIMenu
Private Declare Function GetCapture Lib "User32" () As Long

Private WithEvents oMenuCanvas As PictureBox
Attribute oMenuCanvas.VB_VarHelpID = -1
Private WithEvents oCanvasParent As frmTempForm
Attribute oCanvasParent.VB_VarHelpID = -1
Dim oMenu As New clsMenu
Dim lMenuWidth As Long
Private WithEvents oTimer As Timer
Attribute oTimer.VB_VarHelpID = -1
Dim blnFirstInit As Boolean
Private m_MouseOver As Boolean
Dim oPopupMenu As XPUIMenu
Private m_OwnerMenu As XPUIMenu
Dim dteTimeStart As Date
Event Click(Menu As XPUIMenu, MenuItem As MenuItem)
Private m_ImageList As Object
Private m_ID As String
Private m_Xpos As Long
Private m_Ypos As Long
Private m_Shown As Boolean
Private m_CallingItem As Long
Private m_MenuBorderColor As OLE_COLOR
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

Friend Property Get MenuBorderColor() As OLE_COLOR
    MenuBorderColor = m_MenuBorderColor
End Property

Friend Property Let MenuBorderColor(ByVal Value As OLE_COLOR)
    m_MenuBorderColor = Value
End Property

Friend Property Get CallingItem() As Long
    CallingItem = m_CallingItem
End Property

Friend Property Let CallingItem(ByVal Value As Long)
    m_CallingItem = Value
End Property

Friend Property Get Shown() As Boolean
    Shown = m_Shown
End Property

Friend Property Let Shown(ByVal Value As Boolean)
    m_Shown = Value
End Property

Friend Property Get Ypos() As Long
    Ypos = m_Ypos
End Property

Friend Property Let Ypos(ByVal Value As Long)
    m_Ypos = Value
End Property

Friend Property Get Xpos() As Long
    Xpos = m_Xpos
End Property

Friend Property Let Xpos(ByVal Value As Long)
    m_Xpos = Value
End Property


Friend Property Get ID() As String
    ID = m_ID
End Property

Friend Property Let ID(ByVal Value As String)
    m_ID = Value
End Property

Friend Property Get ImageList() As Object
    Set ImageList = m_ImageList
End Property

Friend Property Set ImageList(ByVal Value As Object)
    Set m_ImageList = Value
End Property

Friend Property Get OwnerMenu() As XPUIMenu
    Set OwnerMenu = m_OwnerMenu
End Property

Friend Property Set OwnerMenu(ByVal Value As XPUIMenu)
    Set m_OwnerMenu = Value
End Property

Friend Property Get MouseOver() As Boolean
    MouseOver = m_MouseOver
End Property

Friend Property Let MouseOver(ByVal Value As Boolean)
    m_MouseOver = Value
End Property

Public Property Get Key() As String
    Key = m_Key
End Property

Public Property Let Key(ByVal Value As String)
    m_Key = Value
End Property


Public Property Get MenuItems() As MenuItems
    If mvarMenuItems Is Nothing Then
        Set mvarMenuItems = New MenuItems
    End If
    Set MenuItems = mvarMenuItems
End Property


Public Property Set MenuItems(vData As MenuItems)
    Set mvarMenuItems = vData
End Property

Private Sub Class_Terminate()
    oMenu.Destroy
    Set mvarMenuItems = Nothing
    Set oMenuCanvas = Nothing
    Set oTimer = Nothing
    Set oCanvasParent = Nothing
End Sub

Public Sub ShowMenu(Xpos As Long, Ypos As Long)
    If oPopupMenu Is Nothing Then
    
    Else
        Set oPopupMenu = Nothing
    End If
    'Set oCanvasParent = Nothing
    Set oCanvasParent = New frmTempForm
    'Set oMenuCanvas = Nothing
    Set oMenuCanvas = oCanvasParent.picCanvas
    
    
    blnFirstInit = False
    Dim lMaxWidth As Long
    Dim lMaxHeight As Long
    Dim RCItemBounds As Rect
    Dim oMenuItem As MenuItem
    
    For Each oMenuItem In Me.MenuItems
        oMenuItem.Popped = False
        oMenuItem.MenuBackColor = m_MenuBackColor
        oMenuItem.MenuImageBackColor = m_MenuImageBackColor
        oMenuItem.MenuItemHotColor = m_MenuItemHotColor
        oMenuItem.separatorcolor = m_separatorcolor
        oMenuItem.MenuItemBorderColor = m_MenuItemBorderColor
        Set oMenuItem.ExpandImage = oCanvasParent.imgExpand.Picture
        DrawText oMenuCanvas.HDC, oMenuItem.Caption, Len(oMenuItem.Caption), RCItemBounds, DT_LEFT Or DT_CALCRECT
        If oMenuItem.Seperator = True Then
            lMaxHeight = lMaxHeight + 5
        Else
            lMaxHeight = CLng(lMaxHeight + (oMenuCanvas.TextHeight("gW") * 1.5))
        End If
        
        If (RCItemBounds.Right - RCItemBounds.Left) > lMaxWidth Then
            lMaxWidth = (RCItemBounds.Right - RCItemBounds.Left)
        End If
    Next
    oMenuCanvas.BackColor = m_MenuBackColor
    oMenuCanvas.Height = (lMaxHeight + 8)
     
    lMenuWidth = (26) + (lMaxWidth + (oMenuCanvas.TextWidth("ABC6")))
    oMenuCanvas.Width = lMenuWidth
    oMenu.MenuBorderColor = Me.MenuBorderColor
    oMenu.Create oMenuCanvas
    pDrawBanner
    
    Set oCanvasParent.oCallerObject = oMenu
    oMenu.Show Xpos, Ypos
    Set oTimer = oCanvasParent.Timer1
    Me.Xpos = Xpos
    Me.Ypos = Ypos
    Me.Shown = True
End Sub


Private Sub oCanvasParent_Load()

End Sub

Private Sub oCanvasParent_MenuClick()
    Dim oMenuItem As MenuItem
    For Each oMenuItem In Me.MenuItems
        If oMenuItem.MouseOver = True Then
            If oMenuItem.Popped = False Then
                If OwnerMenu Is Nothing Then
                    Clicked
                    RaiseEvent Click(Me, oMenuItem)
                Else
                    OwnerMenu.PassClick Me, oMenuItem
                End If
            End If
        End If
    Next
End Sub

Private Sub oMenuCanvas_LostFocus()
    Debug.Print "lOST fOCUS"
End Sub

Private Sub oMenuCanvas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim PT As POINTAPI
    GetCursorPos PT
    Dim blnFoundIT As Boolean
    Dim hw As Long
    hw = WindowFromPoint(PT.x, PT.y)
    If hw <> oMenuCanvas.HWND Then
        
        If oPopupMenu Is Nothing Then
        
        Else
            blnFoundIT = oPopupMenu.IsItMe(hw)
        End If
        If OwnerMenu Is Nothing Then
            
        Else
            blnFoundIT = OwnerMenu.IsItMe(hw)
        End If
        If blnFoundIT = False Then
            Clicked
        End If
    End If
End Sub

Private Sub oMenuCanvas_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Debug.Print "The mouse is moving"
    pCheckPos CLng(x), CLng(y)
    Dim PT As POINTAPI
    GetCursorPos PT
    Dim blnFoundIT As Boolean
    Dim hw As Long
    hw = WindowFromPoint(PT.x, PT.y)
    If hw <> oMenuCanvas.HWND Then
        
        If oPopupMenu Is Nothing Then
        
        Else
            blnFoundIT = oPopupMenu.IsItMe(hw)
        End If
        If OwnerMenu Is Nothing Then
            
        Else
            blnFoundIT = OwnerMenu.IsItMe(hw)
        End If
        If blnFoundIT = False Then
            
        End If
    End If
    
End Sub

Private Sub pDrawBanner()
    Dim oRCBounds As Rect
    oRCBounds.Left = 1
    oRCBounds.Right = oRCBounds.Left + 24
    oRCBounds.Top = 2
    oRCBounds.Bottom = oRCBounds.Top + (oMenuCanvas.Height - 8)
    BoxRect3DDCex oMenuCanvas.HDC, oRCBounds, m_MenuImageBackColor, m_MenuImageBackColor, m_MenuImageBackColor  ' &HDEEDEF, &HDEEDEF, &HDEEDEF
End Sub

Friend Sub pCheckPos(Xpos As Long, Ypos As Long)
    Dim oRCItemRect As Rect
    Dim oMenuItem As MenuItem
    Dim lItemHeight As Long
    Dim iCount As Integer
    Dim iCurrentItem As Integer
    lItemHeight = (oMenuCanvas.TextHeight("gW") * 1.5)
    oRCItemRect.Top = 2
    oRCItemRect.Left = 1
    oRCItemRect.Right = lMenuWidth - 5
    oRCItemRect.Bottom = oRCItemRect.Top
    For Each oMenuItem In Me.MenuItems
        If oMenuItem.Seperator Then
            oRCItemRect.Top = oRCItemRect.Bottom + 2
            oRCItemRect.Bottom = oRCItemRect.Top + 1
            oRCItemRect.Left = 28
            'BoxRect3DDCex oMenuCanvas.HDC, oRCItemRect, &HB8C2C5, &HB8C2C5, &HB8C2C5
            oMenuItem.Top = oRCItemRect.Top
            oMenuItem.Left = oRCItemRect.Left
            oMenuItem.Right = oRCItemRect.Right
            oMenuItem.Bottom = oRCItemRect.Bottom
            Set oMenuItem.MenuObj = oMenuCanvas
            
            oRCItemRect.Bottom = oRCItemRect.Bottom + 2
            oRCItemRect.Left = 1
        Else
            oRCItemRect.Top = oRCItemRect.Bottom
            oRCItemRect.Bottom = oRCItemRect.Top + lItemHeight
            oRCItemRect.Left = 2
            oRCItemRect.Right = lMenuWidth - 6
            If Xpos > oRCItemRect.Left And Xpos < oRCItemRect.Right And Ypos > oRCItemRect.Top And Ypos < oRCItemRect.Bottom Then
                'BoxRect3DDCex oMenuCanvas.HDC, oRCItemRect, vbRed, vbRed, &HFFC0C0
                oMenuItem.Top = oRCItemRect.Top
                oMenuItem.Left = oRCItemRect.Left
                oMenuItem.Right = oRCItemRect.Right
                oMenuItem.Bottom = oRCItemRect.Bottom
                Set oMenuItem.MenuObj = oMenuCanvas
                iCurrentItem = oMenuItem.Item
                If oMenuItem.MouseOver <> True Then
                    oMenuItem.MouseOver = True
                End If
            Else
                oMenuItem.Top = oRCItemRect.Top
                oMenuItem.Left = oRCItemRect.Left
                oMenuItem.Right = oRCItemRect.Right
                oMenuItem.Bottom = oRCItemRect.Bottom
                Set oMenuItem.MenuObj = oMenuCanvas
                oRCItemRect.Right = oRCItemRect.Left + 23
                'BoxRect3DDCex oMenuCanvas.HDC, oRCItemRect, &HDEEDEF, &HDEEDEF, &HDEEDEF
                oRCItemRect.Left = oRCItemRect.Right
                oRCItemRect.Right = lMenuWidth - 6
                'BoxRect3DDCex oMenuCanvas.HDC, oRCItemRect, vbWhite, vbWhite, vbWhite
                oRCItemRect.Left = 2
                If oMenuItem.MouseOver <> False Then
                    oMenuItem.MouseOver = False
                End If
            End If
            oRCItemRect.Left = 1
            oRCItemRect.Right = lMenuWidth - 5
            oRCItemRect.Left = 29
            'DrawText oMenuCanvas.HDC, oMenuItem.Caption, Len(oMenuItem.Caption), oRCItemRect, 564
            oRCItemRect.Left = 1
        End If
    Next
    
End Sub

Private Sub pDrawItems()
    Dim oMenuItem As MenuItem
    For Each oMenuItem In MenuItems
        Set oMenuItem.ImageList = m_ImageList
        oMenuItem.pDrawItem
    Next
'    Dim oRCItemRect As Rect
'    Dim oMenuItem As MenuItem
'    Dim lItemHeight As Long
'    Dim iCount As Integer
'    Dim iCurrentItem As Integer
'    lItemHeight = (oMenuCanvas.TextHeight("gW") * 1.5)
'    oRCItemRect.Top = 2
'    oRCItemRect.Left = 1
'    oRCItemRect.Right = lMenuWidth - 5
'    oRCItemRect.Bottom = oRCItemRect.Top
'    For Each oMenuItem In Me.MenuItems
'        If oMenuItem.Seperator Then
'            oRCItemRect.Top = oRCItemRect.Bottom + 2
'            oRCItemRect.Bottom = oRCItemRect.Top + 1
'            oRCItemRect.Left = 28
'            BoxRect3DDCex oMenuCanvas.HDC, oRCItemRect, &HB8C2C5, &HB8C2C5, &HB8C2C5
'            oRCItemRect.Bottom = oRCItemRect.Bottom + 2
'            oRCItemRect.Left = 1
'        Else
'            oRCItemRect.Top = oRCItemRect.Bottom
'            oRCItemRect.Bottom = oRCItemRect.Top + lItemHeight
'            oRCItemRect.Left = 2
'            oRCItemRect.Right = lMenuWidth - 6
'            If oMenuItem.MouseOver Then
'                BoxRect3DDCex oMenuCanvas.HDC, oRCItemRect, vbRed, vbRed, &HFFC0C0
'                iCurrentItem = oMenuItem.Item
'                oMenuItem.MouseOver = True
'            Else
'                oRCItemRect.Right = oRCItemRect.Left + 23
'                BoxRect3DDCex oMenuCanvas.HDC, oRCItemRect, &HDEEDEF, &HDEEDEF, &HDEEDEF
'                oRCItemRect.Left = oRCItemRect.Right
'                oRCItemRect.Right = lMenuWidth - 6
'                BoxRect3DDCex oMenuCanvas.HDC, oRCItemRect, vbWhite, vbWhite, vbWhite
'                oRCItemRect.Left = 2
'                oMenuItem.MouseOver = False
'            End If
'            oRCItemRect.Left = 1
'            oRCItemRect.Right = lMenuWidth - 5
'            oRCItemRect.Left = 29
'            DrawText oMenuCanvas.HDC, oMenuItem.Caption, Len(oMenuItem.Caption), oRCItemRect, 564
'            oRCItemRect.Left = 1
'        End If
'    Next
'    If iCurrentItem > 0 Then
'        If MenuItems(iCurrentItem).XPUIMenu.MenuItems.Count = 0 Then
'            Set oActivePopup = Nothing
'        Else
'            Dim oWinPos As Rect
'            GetWindowRect oMenuCanvas.HWND, oWinPos
'            Set oActivePopup = Nothing
'            Set oActivePopup = New XPUIMenu
'            Set oActivePopup.MenuItems = Me.MenuItems(iCurrentItem).XPUIMenu.MenuItems
'            oActivePopup.ShowMenu oWinPos.Right - 4, oWinPos.Top + (lItemHeight * iCurrentItem)
'        End If
'    Else
'        Set oActivePopup = Nothing
'    End If

End Sub

Private Sub oTimer_Timer()
    On Error GoTo ErrHandle
    If blnFirstInit = False Then
        pCheckPos 0, 0
        pDrawItems
        blnFirstInit = True
    End If
    
    Dim PT As POINTAPI
    GetCursorPos PT
    
    Dim hw As Long
    hw = WindowFromPoint(PT.x, PT.y)
    
    If hw <> oMenuCanvas.HWND Then
        pCheckPos 0, 0
        m_MouseOver = False
    Else
        m_MouseOver = True
    End If
    Dim oPT As POINTAPI
    ClientToScreen oMenuCanvas.HWND, oPT
    Dim oMenuItems As MenuItem
    For Each oMenuItems In MenuItems
        If oMenuItems.XPUIMenu.MenuItems.Count > 0 And oMenuItems.MouseOver = True And oMenuItems.Disabled = False Then
            If oPopupMenu Is Nothing Then
                Set oPopupMenu = oMenuItems.XPUIMenu
                Set oPopupMenu.ImageList = m_ImageList
                Set oPopupMenu.OwnerMenu = Me
                oPopupMenu.CallingItem = oMenuItems.Item
                oPopupMenu.MenuBorderColor = m_MenuBorderColor
                oPopupMenu.MenuBackColor = m_MenuBackColor
                oPopupMenu.MenuImageBackColor = m_MenuImageBackColor
                oPopupMenu.separatorcolor = m_separatorcolor
                oPopupMenu.MenuItemHotColor = m_MenuItemHotColor
                oPopupMenu.MenuItemBorderColor = m_MenuItemBorderColor
                oMenuItems.Popped = True
                oPopupMenu.ShowMenu oPT.x + ((oMenuItems.Right - oMenuItems.Left) + 3), oPT.y + oMenuItems.Top
                
            Else
                
                If (oPopupMenu.Xpos <> oPT.x + ((oMenuItems.Right - oMenuItems.Left) + 3) Or oPopupMenu.Ypos <> oPT.y + oMenuItems.Top) Then
                    oPopupMenu.Hide
                    MenuItems(oPopupMenu.CallingItem).Popped = False
                    MenuItems(oPopupMenu.CallingItem).pDrawItem
                    DoEvents
                    Set oPopupMenu = Nothing
                    Set oPopupMenu = oMenuItems.XPUIMenu
                    Set oPopupMenu.ImageList = m_ImageList
                    oPopupMenu.MenuBorderColor = m_MenuBorderColor
                    oPopupMenu.MenuImageBackColor = m_MenuImageBackColor
                    oPopupMenu.MenuItemHotColor = m_MenuItemHotColor
                    oPopupMenu.separatorcolor = m_separatorcolor
                    oPopupMenu.MenuItemBorderColor = m_MenuItemBorderColor
                    Set oPopupMenu.OwnerMenu = Me
                    oPopupMenu.MenuBackColor = m_MenuBackColor
                    oMenuItems.Popped = True
                    'oMenuItems.pDrawItem
                    oPopupMenu.CallingItem = oMenuItems.Item
                    oPopupMenu.ShowMenu oPT.x + ((oMenuItems.Right - oMenuItems.Left) + 3), oPT.y + oMenuItems.Top
                Else
                
                End If
            End If
        ElseIf oMenuItems.XPUIMenu.MenuItems.Count > 0 And oMenuItems.MouseOver = False And oMenuItems.Disabled = False Then
            'If (oPopupMenu.Xpos <> oPT.x + ((oMenuItems.Right - oMenuItems.Left) + 3) Or oPopupMenu.Ypos <> oPT.Y + oMenuItems.Top) Then
                'oMenuItems.Popped = False
            'Else
                If oMenuItems.XPUIMenu.Shown Then
                    oMenuItems.Popped = True
                Else
                    oMenuItems.Popped = False
                End If
            'End If
        ElseIf oMenuItems.MouseOver = True Then
            If oPopupMenu Is Nothing Then
            
            Else
                
                oPopupMenu.Hide
                MenuItems(oPopupMenu.CallingItem).Popped = False
                MenuItems(oPopupMenu.CallingItem).pDrawItem
                Set oPopupMenu = Nothing
            End If
            
        Else
            oMenuItems.Popped = False
            'oMenuItems.pDrawItem
        End If
    Next
'    If GetAsyncKeyState(&H1) Then
'        Term
'    End If

Exit Sub
ErrHandle:


End Sub

Friend Sub Term()
    If OwnerMenu Is Nothing Then
        oMenu.Destroy
        
        Set oPopupMenu = Nothing
        Set oMenu = Nothing
        Set oMenuCanvas = Nothing
        Set oCanvasParent = Nothing
    Else
        oMenu.Destroy
        OwnerMenu.Term
        
    End If

End Sub

Friend Function IsItMe(Handle As Long) As Boolean
    On Error Resume Next
    If Handle = oMenuCanvas.HWND Then
        IsItMe = True
        SetCapture oMenuCanvas.HWND
    Else
        If OwnerMenu Is Nothing Then ' I must be first menu
            IsItMe = False
        Else
            IsItMe = OwnerMenu.IsItMe(Handle)
        End If
    End If
End Function

Friend Sub Hide()
    If oPopupMenu Is Nothing Then
    
    Else
        oPopupMenu.Hide
    End If
    Me.Shown = False
    oMenu.Destroy
End Sub

Friend Sub Clicked()
    If oPopupMenu Is Nothing Then
        Term
    Else
        oPopupMenu.Clicked
    End If
End Sub

Friend Sub PassClick(Menu As XPUIMenu, MenuItem As MenuItem)
    If OwnerMenu Is Nothing Then
        
        Clicked
        RaiseEvent Click(Menu, MenuItem)
    Else
        OwnerMenu.PassClick Menu, MenuItem
    End If
End Sub
