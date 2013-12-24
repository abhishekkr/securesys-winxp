VERSION 5.00
Begin VB.Form ieset03 
   BorderStyle     =   0  'None
   Caption         =   "Internet Explorer Settings"
   ClientHeight    =   4785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7275
   Icon            =   "ieset03.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4785
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   120
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   0
      Picture         =   "ieset03.frx":1159A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5880
   End
   Begin VB.Image Image2 
      Height          =   195
      Left            =   5880
      Picture         =   "ieset03.frx":1203E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   6240
      Picture         =   "ieset03.frx":120A2
      Stretch         =   -1  'True
      ToolTipText     =   "ALTER Always On Top"
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   5880
      Picture         =   "ieset03.frx":12106
      Top             =   0
      Width           =   1410
   End
End
Attribute VB_Name = "ieset03"
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
ieset03a.SetFocus
End Sub

Private Sub Form_Paint()
ieset03a.Left = ieset03.Left + 250
ieset03a.Top = ieset03.Top + 350
ieset03a.Show
ieset03a.SetFocus
End Sub

Private Sub Image1_Click()
effect = 2
loader.t1.Text = "1"
Timer1.Enabled = True
End Sub

Private Sub Image2_Click()
Me.WindowState = 1
ieset03a.Hide
End Sub

Private Sub Image3_Click()
StayOnTop
End Sub

Private Sub Image4_Click()
ieset03a.SetFocus
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ieset03a.SetFocus
End Sub

Private Sub Timer1_Timer()
If effect = 2 Then
  trans = trans - 10
 MakeTransparent Me.hwnd, trans
 If trans < 10 Then
  effect = 1
  Unload Me
  Unload ieset03a
 End If
ElseIf effect = 1 Then
 trans = trans + 10
 MakeTransparent Me.hwnd, trans
 If trans > 180 Then
  Timer1.Enabled = False
 End If
End If
End Sub



