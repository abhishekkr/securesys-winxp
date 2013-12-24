VERSION 5.00
Begin VB.Form frmset 
   BorderStyle     =   0  'None
   Caption         =   "Settings 'n Help"
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6990
   Icon            =   "frmset.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   480
      Top             =   120
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Always On Top"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Windows Up Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      ToolTipText     =   "It shows the time spent after Windows has been been logged in."
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Developer's Website"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "About It"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Help File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Move these tabs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   195
      Left            =   5640
      Picture         =   "frmset.frx":1159A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   6000
      Picture         =   "frmset.frx":115FE
      Stretch         =   -1  'True
      ToolTipText     =   "ALTER Always On Top"
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   5640
      Picture         =   "frmset.frx":11662
      Top             =   0
      Width           =   1410
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   0
      Picture         =   "frmset.frx":11EC0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "frmset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim effect As Integer, trans As Integer, tim As Integer
'Sub: StayOnTop

Public Sub chCcap()
  Dim h As Integer, m As Integer, s As Integer
  s = GWinUp
  h = Int(s / 3600)
  m = Int((s - (h * 3600)) / 60)
  s = Int((s - (h * 3600) - (m * 60)))
  Command5.Caption = h & " hr " & m & " min " & s & " sec"
End Sub

Private Sub Command1_Click()
loader.t1.Text = "1"
Form1.Show
frmset.Hide
End Sub

Private Sub Command2_Click()
MsgBox "To be made", vbInformation, "Work Due"
End Sub

Private Sub Command3_Click()
MsgBox "'ABK Secure Sys v1.1' is the O.S. tweaking software made by" _
& vbCrLf & "Abhishek Kumar [B.C.A.-2003 batch] as my final sem. project.", vbInformation, "about"
End Sub

Private Sub Command4_Click()
MsgBox "Browse 'http://abhikumar163.googlepages.com'" _
& vbCrLf & "for all version downloads, info 'n help on this software.", vbInformation, "Visit:"
End Sub

Private Sub Command5_Click()
  If tim = 0 Then
  tim = 1
  chCcap
  Else
  tim = 0
  Command5.Caption = "Windows Up Time"
  End If
End Sub

Private Sub Command6_Click()
StayOnTop
End Sub

Private Sub Form_Load()
effect = 1
trans = 0
MakeTransparent Me.hwnd, trans
Timer1.Enabled = True
End Sub

Private Sub Image1_Click()
effect = 2
Timer1.Enabled = True
loader.t1.Text = "1"
End Sub

Private Sub Image2_Click()
Me.WindowState = 1
End Sub

Private Sub Image3_Click()
StayOnTop
End Sub

Private Sub Timer1_Timer()
If effect = 2 Then
  trans = trans - 10
 MakeTransparent Me.hwnd, trans
 If trans < 10 Then
  effect = 1
  Unload Me
 End If
ElseIf effect = 1 Then
 trans = trans + 10
 MakeTransparent Me.hwnd, trans
 If trans > 190 Then
  Timer1.Enabled = False
 End If
End If
End Sub



Private Sub Timer2_Timer()
If tim = 1 Then
chCcap
End If
End Sub
