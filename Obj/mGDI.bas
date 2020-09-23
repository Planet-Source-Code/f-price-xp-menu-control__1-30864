Attribute VB_Name = "mGDI"
Option Explicit
DefInt A-Z

Type Rect
 Left       As Long
 Top        As Long
 Right      As Long
 Bottom     As Long
End Type

Type POINTAPI
 x As Long
 Y As Long
End Type
Public Const DT_WORD_ELLIPSIS = &H40000
Public Declare Function PtInRect Lib "User32" (lpRect As Rect, ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function ScreenToClient Lib "User32" (ByVal HWND As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetDC Lib "User32" (ByVal HWND As Long) As Long
Public Declare Function GetWindowRect Lib "User32" (ByVal HWND As Long, lpRect As Rect) As Long
Public Declare Function GetWindowDC Lib "User32" (ByVal HWND As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal HDC As Long, ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal HDC As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function ReleaseDC Lib "User32" (ByVal HWND As Long, ByVal HDC As Long) As Long
Public Declare Function GetDesktopWindow Lib "User32" () As Long

Public Declare Function WindowFromPoint Lib "User32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal HWND As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Public Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long

Public Declare Function GetClientRect Lib "User32" (ByVal HWND As Long, lpRect As Rect) As Long
Public Declare Function GetAsyncKeyState Lib "User32" (ByVal vKey As Long) As Integer
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function CreatePen& Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long)
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function DeleteObject& Lib "gdi32" (ByVal hObject As Long)
Declare Function DrawEdge Lib "User32" (ByVal HDC As Long, qrc As Rect, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Declare Function DrawFocusRect& Lib "User32" (ByVal HDC As Long, lpRect As Rect)
Declare Function DrawFrameControl Lib "User32" (ByVal HDC&, lpRect As Rect, ByVal un1 As Long, ByVal un2 As Long) As Boolean
Declare Function DrawText& Lib "User32" Alias "DrawTextA" (ByVal HDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As Rect, ByVal wFormat As Long)
Declare Function FillRect& Lib "User32" (ByVal HDC As Long, lpRect As Rect, ByVal hBrush As Long)
Declare Function GetBkColor& Lib "gdi32" (ByVal HDC As Long)
Declare Function GetDeviceCaps Lib "gdi32" (ByVal HDC As Long, ByVal nIndex As Long) As Long
Declare Function GetTextColor& Lib "gdi32" (ByVal HDC As Long)
Declare Function LineTo& Lib "gdi32" (ByVal HDC As Long, ByVal x As Long, ByVal Y As Long)
Declare Function MoveToEx& Lib "gdi32" (ByVal HDC As Long, ByVal x As Long, ByVal Y As Long, lpPoint As POINTAPI)
Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Declare Function SelectObject& Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long)
Declare Function SetTextColor& Lib "gdi32" (ByVal HDC As Long, ByVal crColor As Long)
Declare Function SetTextJustification Lib "gdi32" (ByVal HDC As Long, ByVal nBreakExtra As Long, ByVal nBreakCount As Long) As Long
Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal HDC As Long, ByVal x As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Declare Function UpdateWindow& Lib "User32" (ByVal HWND As Long)
Declare Function DrawIconEx Lib "User32" (ByVal HDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

'  flags for DrawFrameControl
Public Const DFC_CAPTION = 1 'Title bar
Public Const DFC_MENU = 2   'Menu
Public Const DFC_SCROLL = 3 'Scroll bar
Public Const DFC_BUTTON = 4 'Standard button

Public Const DFCS_CAPTIONCLOSE = &H0    'Close button
Public Const DFCS_CAPTIONMIN = &H1 'Minimize button
Public Const DFCS_CAPTIONMAX = &H2 'Maximize button
Public Const DFCS_CAPTIONRESTORE = &H3  'Restore button
Public Const DFCS_CAPTIONHELP = &H4     'Windows 95 only: Help button

Public Const DFCS_MENUARROW = &H0 'Submenu arrow
Public Const DFCS_MENUCHECK = &H1 'Check mark
Public Const DFCS_MENUBULLET = &H2 'Bullet
Public Const DFCS_MENUARROWRIGHT = &H4

Public Const DFCS_SCROLLUP = &H0   'Up arrow of scroll bar
Public Const DFCS_SCROLLDOWN = &H1 'Down arrow of scroll bar
Public Const DFCS_SCROLLLEFT = &H2 'Left arrow of scroll bar
Public Const DFCS_SCROLLRIGHT = &H3 'Right arrow of scroll bar

Public Const DFCS_SCROLLCOMBOBOX = &H5   'Combo box scroll bar
Public Const DFCS_SCROLLSIZEGRIP = &H8   'Size grip
Public Const DFCS_SCROLLSIZEGRIPRIGHT = &H10   'Size grip in bottom-right corner of window

Public Const DFCS_BUTTONCHECK = &H0 'Check box
Public Const DFCS_BUTTONRADIO = &H4 'Radio button
Public Const DFCS_BUTTON3STATE = &H8 'Three-state button
Public Const DFCS_BUTTONPUSH = &H10 'Push button
Public Const DFCS_INACTIVE = &H100 'Button is inactive (grayed)
Public Const DFCS_PUSHED = &H200  'Button is pushed
Public Const DFCS_CHECKED = &H400 'Button is checked
Public Const DFCS_ADJUSTRECT = &H2000   'Bounding rectangle is adjusted to exclude the surrounding edge of the push button
Public Const DFCS_FLAT = &H4000   'Button has a flat border
Public Const DFCS_MONO = &H8000   'Button has a monochrome border

Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENOUTER = &H2
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_OUTER = &H3
Public Const BDR_INNER = &HC
Public Const BDR_RAISED = &H5
Public Const BDR_SUNKEN = &HA

Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)

Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_DIAGONAL = &H10

' For diagonal lines, the BF_RECT flags specify the end point of
' the vector bounded by the rectangle parameter.
Public Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)

Public Const BF_MIDDLE = &H800    ' Fill in the middle.
Public Const BF_SOFT = &H1000     ' Use for softer buttons.
Public Const BF_ADJUST = &H2000   ' Calculate the space left over.
Public Const BF_FLAT = &H4000     ' For flat rather than 3-D borders.
Public Const BF_MONO = &H8000     ' For monochrome borders.

'DrawText Constants
Public Const DT_BOTTOM = &H8
Public Const DT_CALCRECT = &H400
Public Const DT_CENTER = &H1
Public Const DT_LEFT = &H0
Public Const DT_NOCLIP = &H100
Public Const DT_NOPREFIX = &H800
Public Const DT_RIGHT = &H2
Public Const DT_SINGLELINE = &H20
Public Const DT_TOP = &H0
Public Const DT_VCENTER = &H4
Public Const DT_WORDBREAK = &H10

Public PT As POINTAPI

Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal Height As Long, ByVal Width As Long, _
    ByVal Escapement As Long, ByVal Orientation As Long, ByVal Weight As Long, ByVal Italic As Long, ByVal Underline As Long, ByVal StrikeOut As _
    Long, ByVal CharSet As Long, ByVal OutputPrecision As Long, ByVal ClipPrecision As Long, ByVal Quality As Long, ByVal PitchAndFamily As _
    Long, ByVal Face As String) As Long
    
Public Const FW_BOLD = 700
Public Const FW_NORMAL = 400
Public Const DEFAULT_CHARSET = 1
Public Const OUT_DEFAULT_PRECIS = 0
Public Const CLIP_DEFAULT_PRECIS = 0
Public Const PROOF_QUALITY = 2
Public Const DEFAULT_PITCH = 0
Public Const FF_DONTCARE = 0
Global blnVerty As Boolean
Global btnCaption As String
Global btnWidth As Integer
Global btnKey As String

Public Type TrackMouseEvent
    cbSize As Long
    dwFlags As Long
    HWND As Long
    dwHoverTime As Long
End Type

Public Const WM_MOUSELEAVE = &H2A3
Public Const TME_LEAVE = &H2

Public Declare Function TrackMouseEvent Lib "comctl32.dll" Alias "_TrackMouseEvent" ( _
    ByRef lpEventTrack As TrackMouseEvent) As Long
    
    Public Declare Function ClientToScreen Lib "User32" (ByVal HWND As Long, lpPoint As POINTAPI) As Long
Public Declare Function SetParent Lib "User32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long


Public Declare Function SetFocusAPI Lib "User32" Alias "SetFocus" (ByVal HWND As Long) As Long
Public Declare Function SetCapture Lib "User32" (ByVal HWND As Long) As Long
Public Declare Function ReleaseCapture Lib "User32" () As Long
Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal HWND As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal HWND As Long, ByVal nIndex As Long) As Long
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_APPWINDOW = &H40000
Public Const WS_EX_TOOLWINDOW = &H80&
Public Declare Function SetWindowPos Lib "User32" (ByVal HWND As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED


Public Declare Function MoveWindow Lib "User32" (ByVal HWND As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Public Declare Function SystemParametersInfo Lib "User32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_GETWORKAREA = 48&
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

' Subclassing object to catch Alt-Tab

Public Const WM_ACTIVATE = &H6
Public Const WM_KEYDOWN = &H100

Public Sub DrawCtlEdge(HDC As Long, x As Single, Y As Single, W As Single, H As Single, Optional Style As Long = EDGE_RAISED, Optional Flags As Long = BF_RECT)
 Dim R As Rect
 With R
  .Left = x
  .Top = Y
  .Right = x + W
  .Bottom = Y + H
 End With
 DrawEdge HDC, R, Style, Flags
End Sub

Public Function DrawControl(ByVal HDC As Long, ByVal x As Single, ByVal Y As Single, ByVal W As Single, ByVal H As Single, ByVal CtlType As Long, ByVal Flags As Long)
 Dim R As Rect
 With R
  .Left = x
  .Top = Y
  .Right = x + W
  .Bottom = Y + H
 End With
 DrawControl = DrawFrameControl(HDC, R, CtlType, Flags)
End Function

Function TranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long = 0) As Long
 If OleTranslateColor(clr, hPal, TranslateColor) Then TranslateColor = -1
End Function
Public Function LineDC(ByVal HDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, Optional Color As OLE_COLOR = -1) As Long
 Dim hPen As Long, hPenOld As Long
 Dim R
 hPen = CreatePen(0, 1, IIf(Color = -1, GetTextColor(HDC), TranslateColor(Color)))
 hPenOld = SelectObject(HDC, hPen)
 MoveToEx HDC, X1, Y1, PT
 LineDC = LineTo(HDC, X2, Y2)
 SelectObject HDC, hPenOld
 DeleteObject hPen
 DeleteObject hPenOld
End Function

Public Sub Box3DDC(ByVal HDC As Long, ByVal x As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, Optional Highlight As OLE_COLOR = vb3DHighlight, Optional Shadow As OLE_COLOR = vb3DShadow, Optional Fill As OLE_COLOR = -1)
 Dim hPen As Long, hPenOld As Long
 'Fill
 If Fill <> -1 Then BoxSolidDC HDC, x, Y, W, H, Fill
 'Highlight
 hPen = CreatePen(0, 1, TranslateColor(Highlight))
 hPenOld = SelectObject(HDC, hPen)
 MoveToEx HDC, x + W - 1, Y, PT
 LineTo HDC, x, Y
 LineTo HDC, x, Y + H - 1
 SelectObject HDC, hPenOld
 DeleteObject hPen
 DeleteObject hPenOld
 'Shadow
 hPen = CreatePen(0, 1, TranslateColor(Shadow))
 hPenOld = SelectObject(HDC, hPen)
 LineTo HDC, x + W - 1, Y + H - 1
 LineTo HDC, x + W - 1, Y
 SelectObject HDC, hPenOld
 DeleteObject hPen
 DeleteObject hPenOld
End Sub
Public Sub BoxDC(ByVal HDC As Long, ByVal x As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, Optional Color As OLE_COLOR = vbButtonFace, Optional Fill As OLE_COLOR = -1)
 Dim hPen As Long, hPenOld As Long
 'Fill
 If Fill <> -1 Then BoxSolidDC HDC, x, Y, W, H, Fill
 'Box
 hPen = CreatePen(0, 1, TranslateColor(Color))
 hPenOld = SelectObject(HDC, hPen)
 MoveToEx HDC, x + W - 1, Y, PT
 LineTo HDC, x, Y
 LineTo HDC, x, Y + H - 1
 LineTo HDC, x + W - 1, Y + H - 1
 LineTo HDC, x + W - 1, Y
 SelectObject HDC, hPenOld
 DeleteObject hPen
 DeleteObject hPenOld
End Sub

Public Function BoxSolidDC(ByVal HDC As Long, ByVal x As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, Optional ByVal Fill As OLE_COLOR = vbButtonFace)
 Dim hBrush As Long
 Dim R As Rect
 hBrush = CreateSolidBrush(TranslateColor(Fill))
 With R
  .Left = x
  .Top = Y
  .Right = x + W - 1
  .Bottom = Y + H - 1
 End With
 FillRect HDC, R, hBrush
 DeleteObject hBrush
End Function

Public Sub BoxRect3DDC(ByVal HDC As Long, R As Rect, Optional Highlight As OLE_COLOR = vb3DHighlight, Optional Shadow As OLE_COLOR = vb3DShadow, Optional Fill As OLE_COLOR = -1)
 Box3DDC HDC, R.Left, R.Top, R.Right - R.Left, R.Bottom - R.Top, Highlight, Shadow, Fill
End Sub

Public Sub PaintText(ByVal HDC As Long, ByVal Text$, ByVal x As Single, ByVal Y As Single, ByVal W As Single, ByVal H As Single, Optional ByVal Flags As Long = DT_LEFT)
 Dim R As Rect
 With R
  .Left = x
  .Top = Y
  .Right = x + W
  .Bottom = Y + H
 End With
 DrawText HDC, Text$, -1, R, Flags
End Sub


Public Sub DrawFocus(ByVal HDC As Long, ByVal x As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long)
 Dim R As Rect
 With R
  .Left = x
  .Top = Y
  .Right = x + W
  .Bottom = Y + H
 End With
 DrawFocusRect HDC, R
 
End Sub


Public Function MouseUnder(HWND As Long) As Boolean
    Dim ptMouse As POINTAPI
    GetCursorPos ptMouse
    If WindowFromPoint(ptMouse.x, ptMouse.Y) = HWND Then
       MouseUnder = True
    Else
       MouseUnder = False
    End If

End Function





Public Sub SetPosition(frm As Form, HWND As Long)

Dim rc As Rect
Dim SX As Long
Dim SY As Long

    SY = Screen.TwipsPerPixelY
    SX = Screen.TwipsPerPixelX
    
    GetWindowRect HWND, rc
    If frm.Height + (rc.Bottom + 1) * SY > Screen.Height Then
        rc.Top = (rc.Top - 1) * SY - frm.Height
    Else
        rc.Top = (rc.Bottom + 1) * SY
    End If
    
    If frm.Width + rc.Left * SX > Screen.Width Then
        rc.Left = Screen.Width - frm.Width - SX
    Else
        rc.Left = rc.Left * SX
    End If
    
    frm.Move rc.Left, rc.Top
   
End Sub

Public Function InflateRect(Rect As Rect, Value As Integer) As Integer
    Rect.Top = Rect.Top - Value
    Rect.Left = Rect.Left - Value
    Rect.Bottom = Rect.Bottom + Value
    Rect.Right = Rect.Right + Value
End Function

Public Function DeflateRect(Rect As Rect, Value As Integer) As Integer
    Rect.Top = Rect.Top + Value
    Rect.Left = Rect.Left + Value
    Rect.Bottom = Rect.Bottom - Value
    Rect.Right = Rect.Right - Value
End Function

Public Sub zCopyDC(ByVal lHDCDest As Long, ByVal lHDCSource As Long, ByRef tR As Rect, Top As Integer, Left As Integer)
    With tR
        BitBlt lHDCDest, .Left, .Top, .Right - .Left, .Bottom - .Top, lHDCSource, .Left, .Top, vbSrcCopy
    End With
End Sub

Public Sub BoxRect3DDCex(ByVal HDC As Long, R As Rect, Optional Highlight As OLE_COLOR = vb3DHighlight, Optional Shadow As OLE_COLOR = vb3DShadow, Optional Fill As OLE_COLOR = -1)
 Box3DDCex HDC, R.Left, R.Top, R.Right - R.Left, R.Bottom - R.Top, Highlight, Shadow, Fill
End Sub

Public Sub Box3DDCex(ByVal HDC As Long, ByVal x As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, Optional Highlight As OLE_COLOR = vb3DHighlight, Optional Shadow As OLE_COLOR = vb3DShadow, Optional Fill As OLE_COLOR = -1)
 Dim hPen As Long, hPenOld As Long
 'Fill
 If Fill <> -1 Then BoxSolidDC HDC, x, Y, W, H, Fill
 'Highlight
 hPen = CreatePen(0, 1, TranslateColor(Highlight))
 hPenOld = SelectObject(HDC, hPen)
 MoveToEx HDC, x + W - 1, Y, PT
 LineTo HDC, x, Y
 LineTo HDC, x, Y + H - 1
 SelectObject HDC, hPenOld
 DeleteObject hPen
 DeleteObject hPenOld
 'Shadow
 hPen = CreatePen(0, 1, TranslateColor(Shadow))
 hPenOld = SelectObject(HDC, hPen)
 MoveToEx HDC, x, Y + H - 1, PT
 LineTo HDC, x + W - 1, Y + H - 1
 LineTo HDC, x + W - 1, Y - 1
 SelectObject HDC, hPenOld
 DeleteObject hPen
 DeleteObject hPenOld
End Sub

Public Sub DrawMenuShadow( _
    ByVal HWND As Long, _
    ByVal HDC As Long, _
    ByVal xOrg As Long, _
    ByVal yOrg As Long)
     
    Dim hDcDsk As Long
    Dim Rec As Rect
    Dim winW As Long, winH As Long
    Dim x As Long, Y As Long, c As Long
     
    '- Get the size of the menu...
    GetWindowRect HWND, Rec
    winW = Rec.Right - Rec.Left
    winH = Rec.Bottom - Rec.Top
     
    ' - Get the desktop hDC...
    hDcDsk = GetWindowDC(GetDesktopWindow)
     
    ' - Simulate a shadow on right edge...
    For x = 1 To 4
        For Y = 0 To 3
            c = GetPixel(hDcDsk, xOrg + winW - x, yOrg + Y)
            SetPixel HDC, winW - x, Y, c
        Next Y
        For Y = 4 To 7
            c = GetPixel(hDcDsk, xOrg + winW - x, yOrg + Y)
            SetPixel HDC, winW - x, Y, pMask(3 * x * (Y - 3), c)
        Next Y
        For Y = 8 To winH - 5
            c = GetPixel(hDcDsk, xOrg + winW - x, yOrg + Y)
            SetPixel HDC, winW - x, Y, pMask(15 * x, c)
        Next Y
        For Y = winH - 4 To winH - 1
            c = GetPixel(hDcDsk, xOrg + winW - x, yOrg + Y)
            SetPixel HDC, winW - x, Y, pMask(3 * x * -(Y - winH), c)
        Next Y
    Next x
     
    ' - Simulate a shadow on the bottom edge...
    For Y = 1 To 4
        For x = 0 To 3
            c = GetPixel(hDcDsk, xOrg + x, yOrg + winH - Y)
            SetPixel HDC, x, winH - Y, c
        Next x
        For x = 4 To 7
            c = GetPixel(hDcDsk, xOrg + x, yOrg + winH - Y)
            SetPixel HDC, x, winH - Y, pMask(3 * (x - 3) * Y, c)
        Next x
        For x = 8 To winW - 5
            c = GetPixel(hDcDsk, xOrg + x, yOrg + winH - Y)
            SetPixel HDC, x, winH - Y, pMask(15 * Y, c)
        Next x
    Next Y
     
    ' - Release the desktop hDC...
    ReleaseDC GetDesktopWindow, hDcDsk

End Sub

' - Function pMask splits a color
' into its RGB components and
' transforms the color using
' a scale 0..255
Private Function pMask( _
    ByVal lScale As Long, _
    ByVal lColor As Long) As Long
     
    Dim R As Byte
    Dim G As Byte
    Dim B As Byte
     
    Long2RGB lColor, R, G, B
     
    R = pTransform(lScale, R)
    G = pTransform(lScale, G)
    B = pTransform(lScale, B)
     
    pMask = RGB(R, G, B)
     
End Function

' - Function pTransform converts
' a RGB subcolor using a scale
' where 0 = 0 and 255 = lScale
Private Function pTransform( _
    ByVal lScale As Long, _
    ByVal lColor As Long) As Long
     
    pTransform = lColor - Int(lColor * lScale / 255)
End Function

Public Sub Long2RGB(LongColor As Long, R As Byte, G As Byte, B As Byte)
    On Error Resume Next
    ' convert to hex using vb's hex function
    '     , then use the hex2rgb function
    Hex2RGB (Hex(LongColor)), R, G, B
    'Debug.Print r, g, b
End Sub

Public Sub Hex2RGB(strHexColor As String, R As Byte, G As Byte, B As Byte)
    Dim HexColor As String
    Dim I As Byte
    On Error Resume Next
    ' make sure the string is 6 characters l
    '     ong
    ' (it may have been given in &H###### fo
    '     rmat, we want ######)
    strHexColor = Right((strHexColor), 6)
    ' however, it may also have been given a
    '     s or #***** format, so add 0's in front


    For I = 1 To (6 - Len(strHexColor))
        HexColor = HexColor & "0"
    Next
    HexColor = HexColor & strHexColor
    ' convert each set of 2 characters into
    '     bytes, using vb's cbyte function
    R = CByte("&H" & Right$(HexColor, 2))
    G = CByte("&H" & Mid$(HexColor, 3, 2))
    B = CByte("&H" & Left$(HexColor, 2))
End Sub



