VERSION 5.00
Begin VB.Form winstart01 
   BorderStyle     =   0  'None
   Caption         =   "Windows StartUp Controller"
   ClientHeight    =   5085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6780
   Icon            =   "winstart01.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   120
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   5760
      Picture         =   "winstart01.frx":1159A
      Stretch         =   -1  'True
      ToolTipText     =   "ALTER Always On Top"
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   195
      Left            =   5400
      Picture         =   "winstart01.frx":115FE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   5400
      Picture         =   "winstart01.frx":11662
      Top             =   0
      Width           =   1410
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   0
      Picture         =   "winstart01.frx":11EC0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5760
   End
End
Attribute VB_Name = "winstart01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim effect As Integer, trans As Integer
Dim myx As Integer, myy As Integer

Private Sub Form_Load()
effect = 1
trans = 0
MakeTransparent Me.hwnd, trans
Timer1.Enabled = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
winstart01a.SetFocus
End Sub

Private Sub Form_Paint()
winstart01a.Left = winstart01.Left + 250
winstart01a.Top = winstart01.Top + 350
winstart01a.Show
winstart01a.SetFocus
End Sub

Private Sub Image1_Click()
effect = 2
Timer1.Enabled = True
loader.t1.Text = "1"
End Sub


Private Sub Image2_Click()
Me.WindowState = 1
winstart01a.Hide
End Sub

Private Sub Image3_Click()
StayOnTop
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
winstart01a.SetFocus
End Sub

Private Sub Timer1_Timer()
If effect = 2 Then
  trans = trans - 5
 MakeTransparent Me.hwnd, trans
 If trans < 10 Then
  effect = 1
  Unload Me
  Unload winstart01a
 End If
ElseIf effect = 1 Then
 trans = trans + 5
 MakeTransparent Me.hwnd, trans
 If trans > 180 Then
  Timer1.Enabled = False
 End If
End If
End Sub
