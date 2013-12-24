VERSION 5.00
Begin VB.Form loader 
   Caption         =   "Loader"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "loader.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox t1 
      Height          =   285
      Left            =   3360
      TabIndex        =   0
      Text            =   "1"
      Top             =   1440
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1200
      Top             =   2040
   End
End
Attribute VB_Name = "loader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
frmSplash.Show (vbModal)
load
StayOnTop
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
unloadall
End Sub
