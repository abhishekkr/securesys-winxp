VERSION 5.00
Begin VB.Form fileb 
   Caption         =   "File Browser"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6030
   ControlBox      =   0   'False
   Icon            =   "fileb.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2655
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   3000
      TabIndex        =   0
      Top             =   600
      Width           =   2895
   End
End
Attribute VB_Name = "fileb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim str1 As String

Private Sub CancelButton_Click()
str1 = ""
winstart01a.addpath = str1
fileb.Hide
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo erdrv
Dir1.Path = Drive1.Drive
erdrv:
End Sub

Private Sub File1_Click()
If Len(Dir1.Path) = 3 Then
str1 = File1.Path & File1.FileName
Else
str1 = File1.Path & "\" & File1.FileName
End If
End Sub

Private Sub Form_Load()
str1 = ""
End Sub

Private Sub OKButton_Click()
If str1 <> "" Then
winstart01a.addpath = str1
sysetn02a.fpath = str1
End If
fileb.Hide
End Sub

