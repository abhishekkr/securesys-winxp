VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "multiinstAnce"
   ClientHeight    =   870
   ClientLeft      =   1890
   ClientTop       =   1710
   ClientWidth     =   1455
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   1455
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error GoTo cmd1
If loader.t1.Text = "1" Then
loader.t1.Text = "0"
If Command1.Caption = "Windows StartUp Controller" Then
winstart01.Show
winstart01.Timer1.Enabled = True
ElseIf Command1.Caption = "Kool System Settings" Then
loader.t1.Text = "0"
sysetn02.Show
sysetn02.Timer1.Enabled = True
ElseIf Command1.Caption = "Internet Explorer Settings" Then
loader.t1.Text = "0"
ieset03.Show
ieset03.Timer1.Enabled = True
ElseIf Command1.Caption = "Security Settings" Then
loader.t1.Text = "0"
secset04.Show
secset04.Timer1.Enabled = True
ElseIf Command1.Caption = "Settings 'n Help" Then
loader.t1.Text = "0"
frmset.Show
frmset.Timer1.Enabled = True
End If
ElseIf Command1.Caption <> "Exit" Then
MsgBox "Already one feature is being used," _
 & vbCrLf & "at a time one feature is allowed only.", vbCritical, "Use Pattern"
End If
If Command1.Caption = "Exit" Then
 chk = MsgBox("Do you really want to exit.", vbYesNo, "Yes Or No")
 If chk = 6 Then
  unloadall
 Else
  MsgBox "You chose not to exit the application", vbOKOnly, "OK"
 End If
End If
cmd1:
End Sub

Private Sub Form_Paint()
Me.Command1.Width = Me.Width
Me.Command1.Height = Me.Height
End Sub

