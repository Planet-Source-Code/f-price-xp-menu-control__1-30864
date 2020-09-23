VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Implements ISubclass
Private m_lPtr As Long
Private m_hWndParent As Long
Private m_hWndForm As Long
Private m_hWndObject As Long
Private m_bShown As Boolean
Private m_MenuBorderColor As OLE_COLOR

Friend Property Get MenuBorderColor() As OLE_COLOR
    MenuBorderColor = m_MenuBorderColor
End Property

Friend Property Let MenuBorderColor(ByVal Value As OLE_COLOR)
    m_MenuBorderColor = Value
End Property

Private Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
Dim oT As Object
   ' Turn the pointer into an illegal, uncounted interface
   CopyMemory oT, lPtr, 4
   ' Do NOT hit the End button here! You will crash!
   ' Assign to legal reference
   Set ObjectFromPtr = oT
   ' Still do NOT hit the End button here! You will still crash!
   ' Destroy the illegal reference
   CopyMemory oT, 0&, 4
   ' OK, hit the End button if you must--you'll probably still crash
   ' but it will not be the uncounted reference...

End Property

Public Property Get DropDownObject() As Object
   Set DropDownObject = ObjectFromPtr(m_lPtr)
End Property

Public Sub Create(ByRef picThis As Object)
   If (m_lPtr <> 0) Then
      Destroy
   End If
   On Error Resume Next
   m_lPtr = ObjPtr(picThis)
   If (m_lPtr <> 0) Then
      With DropDownObject
         m_hWndParent = .Container.HWND
         .BorderStyle = 0
         .Visible = False
      End With
   End If
   If (Err.Number <> 0) Then
      Err.Raise Err.Number, App.EXEName & ".Create", "Invalid object passed to Create"
      m_lPtr = 0
      m_hWndParent = 0
   End If
   
End Sub
Public Sub Destroy()
   If (m_lPtr <> 0) Then
      SetParent DropDownObject.HWND, m_hWndParent
      m_lPtr = 0
   End If
End Sub
Public Sub Show(ByVal x As Long, ByVal y As Long)
Dim tP As POINTAPI
Dim hWndDesktop As Long
Dim lStyle As Long
Dim lhWnd As Long
Dim lParenthWNd As Long
   
   ' Make sure the picture box won't appear in the
   ' task bar by making it into a Tool Window:
   lhWnd = DropDownObject.HWND
   lStyle = GetWindowLong(lhWnd, GWL_EXSTYLE)
   lStyle = lStyle Or WS_EX_TOOLWINDOW
   lStyle = lStyle And Not (WS_EX_APPWINDOW)
   SetWindowLong lhWnd, GWL_EXSTYLE, lStyle
   
   ' Determine where to show it in Screen coordinates:
   tP.x = x '\ Screen.TwipsPerPixelX
   tP.y = y '\ Screen.TwipsPerPixelY
   lParenthWNd = DropDownObject.Parent.HWND
   'ClientToScreen lParenthWNd, tP
   
   ' Make the picture box a child of the desktop (so
   ' it can be fully shown even if it extends beyond
   ' the form boundaries):
   SetParent lhWnd, hWndDesktop
   
   ' Show the form:
   SetWindowPos lhWnd, hWndDesktop, tP.x, tP.y, DropDownObject.Width \ Screen.TwipsPerPixelX, DropDownObject.Height \ Screen.TwipsPerPixelY, SWP_SHOWWINDOW
   
   ' Tell VB it is shown:
   DropDownObject.Visible = True
   DropDownObject.ZOrder
   
   ' Try to set focus:
   SetFocusAPI lhWnd
   
   ' Capture all mouse messages.
   SetCapture lhWnd
   
   ' Start subclassing for Alt-tab
   m_hWndForm = lParenthWNd
   m_hWndObject = lhWnd

   m_bShown = True
   DrawBorder
   Dim oTmpRect As Rect
   GetWindowRect DropDownObject.HWND, oTmpRect
   DrawMenuShadow DropDownObject.HWND, DropDownObject.HDC, oTmpRect.Left, oTmpRect.Top
   
End Sub

Public Sub Hide()
    On Error Resume Next
   ' Stop subclassing for Alt-tab
   If (m_hWndForm <> 0) Then
      'DetachMessage Me, m_hWndForm, WM_ACTIVATE
   End If
   If (m_hWndObject <> 0) Then
      ' Hide the picturebox:
      DropDownObject.Visible = False
   End If
   m_hWndForm = 0
   m_hWndObject = 0
   
   ' Stop capturing mouse messages:
   ReleaseCapture
         
   ' Store a flag saying we're not shown:
   m_bShown = False
   
End Sub
Public Property Get IsShown() As Boolean
   ' Return whether we are shown or not.
   IsShown = m_bShown
End Property
Public Property Get InRect(ByVal x As Single, ByVal y As Single) As Boolean
    On Error Resume Next
Dim tR As Rect
   If (IsShown()) Then
      GetClientRect DropDownObject.HWND, tR
      x = x \ Screen.TwipsPerPixelX
      y = y \ Screen.TwipsPerPixelY
      If (PtInRect(tR, x, y) = 1) Then
         InRect = True
      End If
   End If
End Property
Public Sub Resize(ByVal lNewWidth As Long, ByVal lNewHeight As Long)
Dim tWR As Rect, tSR As Rect
Dim lR As Long

   ' Get the size of the window on screen:
   GetWindowRect DropDownObject.HWND, tWR
   ' Check if it will fit:
   lR = SystemParametersInfo(SPI_GETWORKAREA, 0, tSR, 0)
   If (lR = 0) Then
      ' Call failed - just use standard screen:
      tSR.Left = 0
      tSR.Top = 0
      tSR.Right = Screen.Width \ Screen.TwipsPerPixelX
      tSR.Bottom = Screen.Height \ Screen.TwipsPerPixelY
   End If
   If (tWR.Left + lNewWidth > tSR.Right) Then
      ' too big in x
      lNewWidth = tSR.Right - tWR.Left
   End If
   If (tWR.Top + lNewHeight > tSR.Bottom) Then
      ' too big in y
      lNewHeight = tSR.Bottom - tWR.Top
   End If
   MoveWindow DropDownObject.HWND, tWR.Left, tWR.Top, lNewWidth, lNewHeight, 1

End Sub

Private Sub Class_Terminate()
   ' Clear up
   Destroy
End Sub

Private Sub DrawBorder()
    Dim oRCBounds As Rect
    GetClientRect DropDownObject.HWND, oRCBounds
    oRCBounds.Right = oRCBounds.Right - 4
    oRCBounds.Bottom = oRCBounds.Bottom - 4
    BoxRect3DDCex DropDownObject.HDC, oRCBounds, m_MenuBorderColor, m_MenuBorderColor ' &H808080, &H808080
End Sub
