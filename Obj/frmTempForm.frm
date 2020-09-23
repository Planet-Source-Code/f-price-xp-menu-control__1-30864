VERSION 5.00
Begin VB.Form frmTempForm 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmTempForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3240
      Top             =   840
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   840
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.Image imgExpand 
      Height          =   240
      Left            =   3240
      Picture         =   "frmTempForm.frx":000C
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmTempForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event MenuClick()
Private m_oCallerObject As Object

Public Property Get oCallerObject() As Object
    Set oCallerObject = m_oCallerObject
End Property

Public Property Set oCallerObject(ByVal Value As Object)
    Set m_oCallerObject = Value
End Property

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' Mouse down handling:
   If (m_oCallerObject.IsShown) Then
      ' Drop down window is visible
      If Not (m_oCallerObject.InRect(x, y)) Then
         ' Mouse down outside drop-down area:
         
         m_oCallerObject.Hide
      Else
         ' Mouse down inside the drop down:
         'Draw floyd, X \ Screen.TwipsPerPixelX, Y \ Screen.TwipsPerPixelY, Button
      End If
   End If

End Sub

Private Sub picCanvas_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim xCellHit As Long, yCellHit As Long, bIn As Boolean

   ' Mouse up.  Determine whether mouse up over a cell:
   'Draw m_cDW, X \ Screen.TwipsPerPixelX, Y \ Screen.TwipsPerPixelY, Button, bIn, xCellHit, yCellHit
   ' Hide the drop down:
   'm_oCallerObject.Hide
   RaiseEvent MenuClick
   ' If an item selected, say what it was:
   If (bIn) Then
      MsgBox "Selected table: " & xCellHit & " x " & yCellHit, vbInformation
   End If

End Sub

