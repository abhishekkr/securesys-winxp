VERSION 5.00
Begin VB.Form sysetn02 
   BorderStyle     =   0  'None
   Caption         =   "Kool System Settings"
   ClientHeight    =   5070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7680
   Icon            =   "sysetn02.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   120
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   6360
      Picture         =   "sysetn02.frx":1159A
      Stretch         =   -1  'True
      ToolTipText     =   "Minimize"
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   6720
      Picture         =   "sysetn02.frx":115FE
      Stretch         =   -1  'True
      ToolTipText     =   "ALTER Always On Top"
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   0
      Picture         =   "sysetn02.frx":11662
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6735
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   6360
      Picture         =   "sysetn02.frx":120C9
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   1410
   End
End
Attribute VB_Name = "sysetn02"
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
sysetn02a.SetFocus
End Sub

Private Sub Form_Paint()
sysetn02a.Left = sysetn02.Left + 250
sysetn02a.Top = sysetn02.Top + 350
sysetn02a.Show
sysetn02a.SetFocus

End Sub

Private Sub Image1_Click()
effect = 2
Timer1.Enabled = True
loader.t1.Text = "1"
End Sub

Private Sub Image2_Click()
Me.WindowState = 1
sysetn02a.Hide
End Sub

Private Sub Image3_Click()
StayOnTop
End Sub

Private Sub Image4_Click()
sysetn02a.SetFocus
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sysetn02a.SetFocus
End Sub

Private Sub Timer1_Timer()
If effect = 2 Then
  trans = trans - 10
 MakeTransparent Me.hwnd, trans
 If trans < 10 Then
  effect = 1
  Unload Me
  Unload sysetn02a
 End If
ElseIf effect = 1 Then
 trans = trans + 10
 MakeTransparent Me.hwnd, trans
 If trans > 180 Then
  Timer1.Enabled = False
 End If
End If
End Sub

