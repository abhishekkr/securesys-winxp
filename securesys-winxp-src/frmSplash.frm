VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   600
      Top             =   3840
   End
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   0
      Width           =   7080
      Begin VB.Label Label2 
         Caption         =   "v1.10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   10
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "A software to modd your O.S...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   9
         Top             =   840
         Width           =   3315
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "A software to secure your PC..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   8
         Top             =   480
         Width           =   3270
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "Freeware License"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "ABK SecureSys"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   2640
         TabIndex        =   6
         Top             =   1320
         Width           =   4230
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Windows XP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4965
         TabIndex        =   5
         Top             =   2340
         Width           =   1890
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Secure It"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5850
         TabIndex        =   4
         Top             =   2700
         Width           =   1005
      End
      Begin VB.Label lblWarning 
         Caption         =   "Information : This software is a freeware developed to keep your PC secure from other users."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   3
         Top             =   3720
         Width           =   6855
      End
      Begin VB.Label lblCompany 
         Caption         =   "http://abhikumar163.googlepages.com/"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   2
         Top             =   3270
         Width           =   2775
      End
      Begin VB.Label lblCopyright 
         Caption         =   "abhikumar163@gmail.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   1
         Top             =   3060
         Width           =   1935
      End
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   240
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   795
         Width           =   2055
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   3780
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim a As Integer
Dim effect As Integer, trans As Integer

Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_DblClick()
Unload Me
End Sub

Private Sub Form_Initialize()
Me.Width = 25
Me.Height = 25
Frame1.Visible = False
lblLicenseTo.Visible = False
lblPlatform.Visible = False
Label1.Visible = False
lblCompanyProduct.Visible = False
lblVersion.Visible = False
lblProductName.Visible = False
imgLogo.Visible = False
lblCopyright.Visible = False
lblCompany.Visible = False
Label2.Visible = False
lblWarning.Visible = False
effect = 1
trans = 0
MakeTransparent Me.hwnd, trans
Timer1.Enabled = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Frame1_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
If Me.Width < 7470 Then
Me.Left = Me.Left - 30
Me.Top = Me.Top - 17
Me.Width = Me.Width + 60
Me.Height = Me.Height + 34
Exit Sub
End If
Frame1.Visible = True
a = a + 2
If a < 1200 Then
If a = 100 Then
lblLicenseTo.Visible = True
ElseIf a = 200 Then
lblPlatform.Visible = True
ElseIf a = 300 Then
Label1.Visible = True
ElseIf a = 400 Then
lblCompanyProduct.Visible = True
ElseIf a = 500 Then
lblVersion.Visible = True
ElseIf a = 600 Then
lblProductName.Visible = True
ElseIf a = 700 Then
imgLogo.Visible = True
ElseIf a = 800 Then
lblCopyright.Visible = True
ElseIf a = 900 Then
lblCompany.Visible = True
ElseIf a = 1000 Then
Label2.Visible = True
ElseIf a > 1100 Then
lblWarning.Visible = True
End If
End If
If a < 1600 Then Exit Sub
effect = 2
Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
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
  Timer2.Enabled = False
 End If
End If
End Sub
