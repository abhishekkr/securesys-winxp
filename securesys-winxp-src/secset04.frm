VERSION 5.00
Begin VB.Form secset04 
   BorderStyle     =   0  'None
   Caption         =   "Security Settings"
   ClientHeight    =   4845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6330
   Icon            =   "secset04.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4845
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   120
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   5280
      Picture         =   "secset04.frx":1159A
      Stretch         =   -1  'True
      ToolTipText     =   "ALTER Always On Top"
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   195
      Left            =   4920
      Picture         =   "secset04.frx":115FE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   4920
      Picture         =   "secset04.frx":11662
      Top             =   0
      Width           =   1410
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   120
      Picture         =   "secset04.frx":11EC0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "secset04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim effect As Integer, trans As Integer


Private Sub Form_Load()
effect = 1
trans = 0
MakeTransparent Me.hwnd, trans
Timer1.Enabled = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
secset04a.SetFocus
End Sub

Private Sub Form_Paint()
secset04a.Left = secset04.Left + 250
secset04a.Top = secset04.Top + 350
secset04a.Show
secset04a.SetFocus
End Sub

Private Sub Image1_Click()
effect = 2
Timer1.Enabled = True
loader.t1.Text = "1"
End Sub

Private Sub Image2_Click()
Me.WindowState = 1
secset04a.Hide
End Sub

Private Sub Image3_Click()
StayOnTop
End Sub

Private Sub Image4_Click()
secset04a.SetFocus
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
secset04a.SetFocus
End Sub

Private Sub Timer1_Timer()
If effect = 2 Then
  trans = trans - 10
 MakeTransparent Me.hwnd, trans
 If trans < 10 Then
  effect = 1
  Unload Me
  Unload secset04a
 End If
ElseIf effect = 1 Then
 trans = trans + 10
 MakeTransparent Me.hwnd, trans
 If trans > 180 Then
  Timer1.Enabled = False
 End If
End If
End Sub


