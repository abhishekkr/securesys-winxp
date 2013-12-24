VERSION 5.00
Begin VB.Form winstart01a 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   4470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6255
   Icon            =   "winstart01a.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4470
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      Caption         =   "ENTRIES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   3600
      TabIndex        =   10
      Top             =   720
      Width           =   1335
      Begin VB.CommandButton Command4 
         Caption         =   "RUN ONCE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   14
         ToolTipText     =   "It's list of programs to be run only after nest startup."
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "HKCU RUN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   13
         ToolTipText     =   "It's list of programs to be run at startup of current user."
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "HKLM RUN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   12
         ToolTipText     =   "It's the list of programs to be run at startup of all users."
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start Up Local"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   11
         ToolTipText     =   "It's list of programs added to startup folder of all users."
         Top             =   2520
         Width           =   855
      End
   End
   Begin VB.TextBox RunTXTPath 
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   285
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Path of keys to Run"
      ToolTipText     =   "It shows the path for file to run of Selected item from list"
      Top             =   120
      Width           =   5415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Height          =   3615
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   3375
      Begin VB.FileListBox ListStartFile 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3210
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "This is the link file that is executed at startup"
         Top             =   360
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.ListBox HKLMRunListKey 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3180
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   3480
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Add New Startup"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   4
      ToolTipText     =   "To add any program at startup of current user."
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "start Regedit"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   3
      ToolTipText     =   "This will display all the kinds of StartUp Entries together in the box......."
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Delete Selected Entry"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   2
      ToolTipText     =   "To delete the selected startup entry."
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command9 
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
      Height          =   615
      Left            =   5040
      TabIndex        =   1
      ToolTipText     =   "Help"
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox addpath 
      Height          =   285
      Left            =   4680
      TabIndex        =   0
      Text            =   "add path"
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Path :"
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
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   615
   End
   Begin VB.Label statuslbl 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   4200
      Width           =   6255
   End
End
Attribute VB_Name = "winstart01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim listKey As Integer
Dim strText, strtextprt As String

Private Sub Command6_Click()
On Error GoTo cmd6
If Me.MousePointer = 14 Then
MsgBox "Click this to add an entry in StartUp entry of current users.", vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
fileb.Show (vbModal)
If addpath.Text <> "" Then
naam = "New"
naam = InputBox("Enter name for key", "Key Name", "NewEntry")
b = SaveValue("HKEY_CURRENT_USER", "Software\Microsoft\Windows\CurrentVersion\Run", naam, addpath.Text, "string")
If b = True Then
strtextprt = "The provided program has been added to current user startup at HKCU RUN"
strText = String(30, " ") + strtextprt
End If
End If
Command3_Click
cmd6:
End Sub

Private Sub Command7_Click()
On Error GoTo cmd7
If Me.MousePointer = 14 Then
MsgBox "Click this to enlist all startup entries together.", vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
Dim a
a = Shell("regedit.exe", vbNormalFocus)
strtextprt = Command7.ToolTipText
strText = String(30, " ") + strtextprt
Exit Sub
cmd7:
MsgBox "System not allowing to invoke Registry Editor", vbCritical, "System Settings"
End Sub

Private Sub Command8_Click()
On Error GoTo cmd8
If Me.MousePointer = 14 Then
MsgBox "Click this to delete the selected entry from list.", vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
If HKLMRunListKey.ListIndex >= 0 And ListStartFile.Visible = False Then
 If listKey = 1 Then
  b = DeleteValue("HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", HKLMRunListKey.Text)
  Command2_Click
 ElseIf listKey = 2 Then
  b = DeleteValue("HKEY_CURRENT_USER", "Software\Microsoft\Windows\CurrentVersion\Run", HKLMRunListKey.Text)
  Command3_Click
 ElseIf listKey = 3 Then
  b = DeleteValue("HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce", HKLMRunListKey.Text)
  If b = False Then b = DeleteValue("HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx", HKLMRunListKey.Text)
  Command4_Click
 End If
End If
If listKey = 4 Then
 Kill ("c:\Documents and Settings\All Users\Start Menu\Programs\Startup\" & ListStartFile.FileName)
 Command1_Click
End If
strtextprt = "The selected entry has been removed from the startup entry."
strText = String(30, " ") + strtextprt
cmd8:
End Sub

Private Sub Command9_Click()
Me.MousePointer = 14
End Sub

Private Sub Form_Click()
If Me.MousePointer = 14 Then
MsgBox "Click this Question Arrow over the button to know its function", vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
End Sub

Private Sub ListStartFile_Click()
If Me.MousePointer = 14 Then
MsgBox "It lists the links in StartUp folder of All Users.", vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
End Sub

Private Sub RunTXTPath_Change()
If Me.MousePointer = 14 Then
MsgBox "This field shows the path of selected startup entry from list below.", vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
End Sub

Private Sub statuslbl_Click()
If Me.MousePointer = 14 Then
MsgBox "Its the status bar showing the action performed.", vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
End Sub

Private Sub Timer1_Timer()
    strText = Mid(strText, 2) & Left(strText, 1)
    statuslbl.Caption = "Status : " & strText
End Sub

Private Sub Command1_Click()
On Error GoTo err
If Me.MousePointer = 14 Then
MsgBox "Click this to list the links in StartUp folder of all users.", vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
listKey = 4
ListStartFile.Visible = True
Frame1.Caption = "StartUp [Local]"
ListStartFile.Path = "c:\Documents and Settings\All Users\Start Menu\Programs\Startup"
strtextprt = " It's list of programs added to startup folder of all users."
ListStartFile.Visible = True
strText = String(30, " ") + strtextprt
err:
End Sub

Private Sub Command2_Click()
On Error GoTo cmd2
If Me.MousePointer = 14 Then
MsgBox "Click this to list the entries in StartUp entry of all users.", vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
Dim gn As String
listKey = 1
HKLMRunListKey.Clear
Frame1.Caption = "HKLM Run"
ListStartFile.Visible = False
gn = listRunkeyvalue("HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Windows\CurrentVersion\Run")
strtextprt = " It's the list of programs to be run at startup of all users."
strText = String(30, " ") + strtextprt
cmd2:
End Sub

Private Sub Command3_Click()
On Error GoTo cmd3
If Me.MousePointer = 14 Then
MsgBox "Click this to list the entries in StartUp entry of current users.", vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
Dim gn As String
listKey = 2
ListStartFile.Visible = False
Frame1.Caption = "HKCU Run"
HKLMRunListKey.Clear
gn = listRunkeyvalue("HKEY_CURRENT_USER", "Software\Microsoft\Windows\CurrentVersion\Run")
strtextprt = " It's list of programs to be run at startup of current user."
strText = String(30, " ") + strtextprt
cmd3:
End Sub

Private Sub Command4_Click()
On Error GoTo cmd4
If Me.MousePointer = 14 Then
MsgBox "Click this to list the entries in StartUp entry of all users to run only after next startup.", vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
Dim gn As String
listKey = 3
HKLMRunListKey.Clear
Frame1.Caption = "Run Once"
ListStartFile.Visible = False
gn = listRunkeyvalue("HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce")
gn = listRunkeyvalue("HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx")
strtextprt = " It's list of programs to be run only after nest startup."
strText = String(30, " ") + strtextprt
cmd4:
End Sub

Private Sub Form_Load()
listKey = 0
strtextprt = " It's a utility to keep a control over which program should run at startup."
strText = String(30, " ") + strtextprt
End Sub

Private Sub HKLMRunListKey_Click()
If Me.MousePointer = 14 Then
MsgBox "It displays the list of StartUp entries in RUN folder.", vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
Dim gn As String
If listKey = 1 Then
gn = GetValue("HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", HKLMRunListKey.Text)
ElseIf listKey = 2 Then
gn = GetValue("HKEY_CURRENT_USER", "Software\Microsoft\Windows\CurrentVersion\Run", HKLMRunListKey.Text)
ElseIf listKey = 3 Then
gn = GetValue("HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce", HKLMRunListKey.Text)
If gn = "" Then gn = GetValue("HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx", HKLMRunListKey.Text)
ElseIf listKey = 4 Then

ElseIf listKey = 5 Then

End If
Me.RunTXTPath.Text = gn
End Sub

