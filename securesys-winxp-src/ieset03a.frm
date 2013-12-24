VERSION 5.00
Begin VB.Form ieset03a 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "IE Settings"
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6720
   Icon            =   "ieset03a.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4095
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   0
      Top             =   3720
      Width           =   375
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00808000&
      Caption         =   "ReInstall"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5160
      TabIndex        =   17
      Top             =   2880
      Width           =   1455
      Begin VB.CommandButton Command8 
         Caption         =   "Enable IE Reinstall"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "It enables the user to reinstall Internet Explorer if needed."
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.ListBox listtmp 
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00808000&
      Caption         =   "Title Text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      TabIndex        =   15
      Top             =   2880
      Width           =   1455
      Begin VB.CommandButton Command3 
         Caption         =   "Change IE Title Text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "It changes title text of IE."
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00808000&
      Caption         =   "Frame5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5160
      TabIndex        =   14
      Top             =   1440
      Width           =   1455
      Begin VB.CommandButton Command11 
         Caption         =   "Delete Typed URL"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "It deleted the selected typed URL."
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton Command10 
         Caption         =   "List All Typed URLs"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   " It will list the typed URLs."
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00808000&
      Caption         =   "Fulscreen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5160
      TabIndex        =   11
      Top             =   0
      Width           =   1455
      Begin VB.CommandButton Command5 
         Caption         =   "Always Fulscreen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "It shifts the registry setting of IE to fullscreen..."
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Reset FulScreen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "It resets fullscreen setting of IE back to normal..."
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   3480
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808000&
      Caption         =   "IE Rt. Clk. Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3720
      TabIndex        =   8
      Top             =   1440
      Width           =   1455
      Begin VB.CommandButton Command4 
         Caption         =   "Right Click Menu Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "It lists the later added Rt.Click Menu item to IE..."
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Delete Menu Item"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "It deletes the selected Menu Entry..."
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      Caption         =   "IE Toolbar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3720
      TabIndex        =   5
      Top             =   0
      Width           =   1455
      Begin VB.CommandButton Command1 
         Caption         =   "List IE ToolBars"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "It lists the additional installed IE toolbars..."
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Delete ToolBar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "It deletes the additional installed toolbars..."
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Caption         =   "Item List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   3495
      Begin VB.ListBox ListIEbar 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   3375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Label statuslbl 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   3840
      Width           =   6375
   End
End
Attribute VB_Name = "ieset03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strText, strtextprt As String

Private Sub Command1_Click()
On Error GoTo cmd1
If Me.MousePointer = 14 Then
MsgBox Command1.ToolTipText, vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
ListIEbar.Clear
b = listIEbarvalue("HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Internet Explorer\Toolbar")
b = listIEbarvalue("HKEY_CURRENT_USER", "Software\Microsoft\Internet Explorer\Toolbars")
strtextprt = " It will list the additional installed toolbars."
strText = String(30, " ") + strtextprt
cmd1:
End Sub

Private Sub Command10_Click()
On Error GoTo cmd10
If Me.MousePointer = 14 Then
MsgBox Command10.ToolTipText, vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
ListIEbar.Clear
b = listIEbarvalue("HKEY_CURRENT_USER", "Software\Microsoft\Internet Explorer\TypedURLs")
strtextprt = " It will list the typed URLs."
strText = String(30, " ") + strtextprt
cmd10:
End Sub

Private Sub Command11_Click()
On Error GoTo cmd11
If Me.MousePointer = 14 Then
MsgBox Command11.ToolTipText, vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
If ListIEbar.ListIndex >= 0 Then
 b = DeleteValue("HKEY_CURRENT_USER", "Software\Microsoft\Internet Explorer\TypedURLs", listtmp.List(ListIEbar.ListIndex))
 Command10_Click
 strtextprt = " It deleted the selected typed URL."
strText = String(30, " ") + strtextprt
End If
cmd11:
End Sub

Private Sub Command2_Click()
On Error GoTo cmd2
If Me.MousePointer = 14 Then
MsgBox Command2.ToolTipText, vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
If ListIEbar.ListIndex >= 0 Then
 b = DeleteValue("HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Internet Explorer\Toolbar", listtmp.List(ListIEbar.ListIndex))
 Command1_Click
 strtextprt = " It deleted the selected additional installed toolbars."
strText = String(30, " ") + strtextprt
End If
cmd2:
End Sub

Private Sub Command3_Click()
On Error GoTo cmd3
If Me.MousePointer = 14 Then
MsgBox Command3.ToolTipText, vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
tt = InputBox("Enter any text you want to be displayed in IE title bar", "Change IE title text")
If tt <> "" Then
b = SaveValue("HKEY_CURRENT_USER", "Software\Microsoft\Internet Explorer\Main", "Window Title", "Internet Explorer : " & tt, "String")
strtextprt = " It changed IE title to Internet Explore : " & tt
strText = String(30, " ") + strtextprt
End If
cmd3:
End Sub

Private Sub Command4_Click()
On Error GoTo cmd4
If Me.MousePointer = 14 Then
MsgBox Command4.ToolTipText, vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
ListIEbar.Clear
b = listIEMENUKey("HKEY_CURRENT_USER", "Software\Microsoft\Internet Explorer\MenuExt")
strtextprt = " It will list the right click menu items."
strText = String(30, " ") + strtextprt
cmd4:
End Sub

Private Sub Command5_Click()
On Error GoTo cmd5
If Me.MousePointer = 14 Then
MsgBox Command5.ToolTipText, vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
b = SaveValue("HKEY_CURRENT_USER", "Software\Microsoft\Internet Explorer\Main", "FullScreen", "yes", "String")
b = SaveValue("HKEY_CURRENT_USER", "Console", "FullScreen", 1, "Dword")
strtextprt = " IE will be FulScreen always."
strText = String(30, " ") + strtextprt
cmd5:
End Sub

Private Sub Command6_Click()
On Error GoTo cmd6
If Me.MousePointer = 14 Then
MsgBox Command6.ToolTipText, vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
If ListIEbar.ListIndex >= 0 Then
 b = DeleteKey("HKEY_CURRENT_USER", "Software\Microsoft\Internet Explorer\MenuExt\" & ListIEbar.Text)
 Command4_Click
 strtextprt = " It deleted the selected menu item."
strText = String(30, " ") + strtextprt
End If
cmd6:
End Sub

Private Sub Command7_Click()
Me.MousePointer = 14
End Sub

Private Sub Command8_Click()
On Error GoTo cmd8
If Me.MousePointer = 14 Then
MsgBox Command8.ToolTipText, vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
b = SaveValue("HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Active Setup\Installed Components\{89820200-ECBD-11cf-8B85-00AA005B4383}", "IsInstalled", "0", "Dword")
strtextprt = " It enables the user to reinstall Internet Explorer if needed."
strText = String(30, " ") + strtextprt
MsgBox "Now User can easily install Internet Explorer on this machine.", vbInformation, "Success"
cmd8:
End Sub

Private Sub Command9_Click()
On Error GoTo cmd9
If Me.MousePointer = 14 Then
MsgBox Command9.ToolTipText, vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
b = SaveValue("HKEY_CURRENT_USER", "Software\Microsoft\Internet Explorer\Main", "FullScreen", "no", "String")
b = SaveValue("HKEY_CURRENT_USER", "Console", "FullScreen", 0, "Dword")
strtextprt = " IE's Fullscreen property has been set to normal."
strText = String(30, " ") + strtextprt
cmd9:
End Sub

Private Sub Form_Load()
On Error GoTo frml
strtextprt = " It's a utility to tweak Internet Explorer."
strText = String(30, " ") + strtextprt
ver = GetValue("HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Internet Explorer\Version Vector", "IE")
Frame1.Caption = "IE ver." & ver
frml:
End Sub

Private Sub statuslbl_Click()
If Me.MousePointer = 14 Then
MsgBox "It's status bar showing all activities progressing.", vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
End Sub

Private Sub Timer1_Timer()
    strText = Mid(strText, 2) & Left(strText, 1)
    statuslbl.Caption = "Status : " & strText
End Sub

