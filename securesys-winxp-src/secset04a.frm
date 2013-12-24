VERSION 5.00
Begin VB.Form secset04a 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Security System"
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5910
   Icon            =   "secset04a.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4230
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox u2 
      BackColor       =   &H00FF8080&
      Caption         =   "Current User"
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
      Left            =   1440
      TabIndex        =   21
      Top             =   3600
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox u1 
      BackColor       =   &H00FF8080&
      Caption         =   "All Users"
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
      TabIndex        =   20
      Top             =   3600
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C000&
      Caption         =   "Registry Edit [RegEdit]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3000
      TabIndex        =   17
      Top             =   0
      Width           =   2775
      Begin VB.CommandButton Command14 
         Caption         =   "Enable Registry Editor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1440
         TabIndex        =   19
         ToolTipText     =   "Enables the User to run RegEdit"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Disable Registry Editor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Disable User Previlege to run RegEdit.......Enhances Security"
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "Control Panel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   2775
      Begin VB.CommandButton Command2 
         Caption         =   "Enable Control Panel"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1440
         TabIndex        =   15
         ToolTipText     =   "It will enable CPL on user..."
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Disable Control Panel"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "It will disable CPL on user..."
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      Caption         =   "Task Manager [ctrl_alt_del]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3000
      TabIndex        =   10
      Top             =   2400
      Width           =   2775
      Begin VB.CommandButton Command4 
         Caption         =   "Enable Task Manager ctrl_alt_del"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1440
         TabIndex        =   12
         ToolTipText     =   "It will enable Task Manager and ctrl_alt_del"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Disable Task Manager ctrl_alt_del"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "It will disable Task Manager and ctrl_alt_del..."
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C000&
      Caption         =   "Windows Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   2775
      Begin VB.CommandButton Command5 
         Caption         =   "Disable Windows Help"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "It will disable MSHelp..."
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Enable Windows Help"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1440
         TabIndex        =   8
         ToolTipText     =   "It will enable MSHelp..."
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C000&
      Caption         =   "Universal Plug && Play Service"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3000
      TabIndex        =   4
      Top             =   1200
      Width           =   2775
      Begin VB.CommandButton Command7 
         Caption         =   "Disable Plug&&Play Service"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "It will disablethe universal Plug 'n Play Services..."
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Enable Plug&&Play Service"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1440
         TabIndex        =   5
         ToolTipText     =   "It will enablethe universal Plug 'n Play Services..."
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C000&
      Caption         =   "Tray Item Display"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   2775
      Begin VB.CommandButton Command9 
         Caption         =   "Disable Showing Tray Icon"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "It will hide all sys tray icon from user view..."
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Enable Showing Tray Icon"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1440
         TabIndex        =   2
         ToolTipText     =   "It will show all sys tray icon from user view..."
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command11 
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
      Left            =   5400
      TabIndex        =   0
      Top             =   3840
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   3720
   End
   Begin VB.Label statuslbl 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   ":"
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
      TabIndex        =   16
      Top             =   3960
      Width           =   5415
   End
End
Attribute VB_Name = "secset04a"
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
strtextprt = Command1.ToolTipText
strText = String(30, " ") + strtextprt
If u2.value = 1 Then _
b = SaveValue("HKEY_CURRENT_USER", "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoControlPanel", 1, "Dword")
If u1.value = 1 Then _
b = SaveValue("HKEY_LOCAL_MACHINE", "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoControlPanel", 1, "Dword")
MsgBox "Control Panel has been Disabled", vbInformation, "Success"
cmd1:
End Sub

Private Sub Command10_Click()
On Error GoTo cmd10
If Me.MousePointer = 14 Then
MsgBox Command10.ToolTipText, vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
strtextprt = Command10.ToolTipText
strText = String(30, " ") + strtextprt
If u2.value = 1 Then _
b = SaveValue("HKEY_CURRENT_USER", "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoTrayItemsDisplay", 0, "Dword")
If u1.value = 1 Then _
b = SaveValue("HKEY_LOCAL_MACHINE", "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoTrayItemsDisplay", 0, "Dword")
MsgBox "Tray Icons has been Enabled", vbInformation, "Success"
cmd10:
End Sub

Private Sub Command11_Click()
Me.MousePointer = 14
End Sub

Private Sub Command13_Click()
On Error GoTo cmd13
If Me.MousePointer = 14 Then
MsgBox Command13.ToolTipText, vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
strtextprt = Command13.ToolTipText
strText = String(30, " ") + strtextprt
If u1.value = 1 Then _
MsgBox "This setting can only be applied to Current user"
If u2.value = 1 Then _
b = SaveValue("HKEY_CURRENT_USER", "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\system", "DisableRegistryTools", 1, "Dword")
MsgBox "Registry Editing has been Disabled for Current User", vbInformation, "Success"
cmd13:
End Sub

Private Sub Command14_Click()
On Error GoTo cmd14
If Me.MousePointer = 14 Then
MsgBox Command14.ToolTipText, vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
strtextprt = Command14.ToolTipText
strText = String(30, " ") + strtextprt
If u1.value = 1 Then _
MsgBox "This setting can only be applied to Current user"
If u2.value = 1 Then _
b = SaveValue("HKEY_CURRENT_USER", "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\system", "DisableRegistryTools", 0, "Dword")
MsgBox "Registry Editing has been Enabled for Current User", vbInformation, "Success"
cmd14:
End Sub

Private Sub Command2_Click()
On Error GoTo cmd2
If Me.MousePointer = 14 Then
MsgBox Command2.ToolTipText, vbInformation, "Help!"
strtextprt = Command2.ToolTipText
Me.MousePointer = 0
Exit Sub
End If
strtextprt = Command2.ToolTipText
strText = String(30, " ") + strtextprt
If u2.value = 1 Then _
b = SaveValue("HKEY_CURRENT_USER", "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoControlPanel", 0, "Dword")
If u1.value = 1 Then _
b = SaveValue("HKEY_LOCAL_MACHINE", "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoControlPanel", 0, "Dword")
MsgBox "Control Panel has been Enabled", vbInformation, "Success"
cmd2:
End Sub

Private Sub Command3_Click()
On Error GoTo cmd3
If Me.MousePointer = 14 Then
MsgBox Command3.ToolTipText, vbInformation, "Help!"
strtextprt = Command3.ToolTipText
Me.MousePointer = 0
Exit Sub
End If
strtextprt = Command3.ToolTipText
strText = String(30, " ") + strtextprt
If u2.value = 1 Then _
b = SaveValue("HKEY_CURRENT_USER", "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", 1, "Dword")
If u1.value = 1 Then _
b = SaveValue("HKEY_LOCAL_MACHINE", "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", 1, "Dword")
MsgBox "Task Manager [Ctrl+Alt+Del] has been Disabled", vbInformation, "Success"
cmd3:
End Sub

Private Sub Command4_Click()
On Error GoTo cmd4
If Me.MousePointer = 14 Then
MsgBox Command4.ToolTipText, vbInformation, "Help!"
strtextprt = Command4.ToolTipText
Me.MousePointer = 0
Exit Sub
End If
strtextprt = Command4.ToolTipText
strText = String(30, " ") + strtextprt
If u2.value = 1 Then _
b = SaveValue("HKEY_CURRENT_USER", "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", 0, "Dword")
If u1.value = 1 Then _
b = SaveValue("HKEY_LOCAL_MACHINE", "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", 0, "Dword")
MsgBox "Task Manager [Ctrl+Alt+Del] has been Enabled", vbInformation, "Success"
cmd4:
End Sub

Private Sub Command5_Click()
On Error GoTo cmd5
If Me.MousePointer = 14 Then
MsgBox Command5.ToolTipText, vbInformation, "Help!"
strtextprt = Command5.ToolTipText
Me.MousePointer = 0
Exit Sub
End If
strText = String(30, " ") + strtextprt
strtextprt = Command5.ToolTipText
If u2.value = 1 Then _
b = SaveValue("HKEY_CURRENT_USER", "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSMHelp", 1, "Dword")
If u1.value = 1 Then _
b = SaveValue("HKEY_LOCAL_MACHINE", "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSMHelp", 1, "Dword")
MsgBox "Windows Help has been Disabled", vbInformation, "Success"
cmd5:
End Sub

Private Sub Command6_Click()
On Error GoTo cmd6
If Me.MousePointer = 14 Then
MsgBox Command6.ToolTipText, vbInformation, "Help!"
strtextprt = Command6.ToolTipText
Me.MousePointer = 0
Exit Sub
End If
strText = String(30, " ") + strtextprt
strtextprt = Command6.ToolTipText
If u2.value = 1 Then _
b = SaveValue("HKEY_CURRENT_USER", "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSMHelp", 0, "Dword")
If u1.value = 1 Then _
b = SaveValue("HKEY_LOCAL_MACHINE", "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSMHelp", 0, "Dword")
MsgBox "Windows Help has been Enabled", vbInformation, "Success"
cmd6:
End Sub

Private Sub Command7_Click()
On Error GoTo cmd7
If Me.MousePointer = 14 Then
MsgBox Command7.ToolTipText, vbInformation, "Help!"
strText = String(30, " ") + strtextprt
strtextprt = Command7.ToolTipText
Me.MousePointer = 0
Exit Sub
End If
strtextprt = Command7.ToolTipText
strText = String(30, " ") + strtextprt
If u1.value = 1 Then
b = SaveValue("HKEY_LOCAL_MACHINE", "SYSTEM\CurrentControlSet\Services\upnphost", "Start", 4, "Dword")
Else
MsgBox "This property can only be applied to All Users", vbInformation
End If
MsgBox "Plug 'n Play Devices has been Disabled", vbInformation, "Success"
cmd7:
End Sub

Private Sub Command8_Click()
On Error GoTo cmd8
If Me.MousePointer = 14 Then
MsgBox Command8.ToolTipText, vbInformation, "Help!"
strtextprt = Command8.ToolTipText
Me.MousePointer = 0
Exit Sub
End If
strtextprt = Command8.ToolTipText
strText = String(30, " ") + strtextprt
If u1.value = 1 Then
b = SaveValue("HKEY_LOCAL_MACHINE", "SYSTEM\CurrentControlSet\Services\upnphost", "Start", 3, "Dword")
Else
MsgBox "This property can only be applied to All Users", vbInformation
End If
MsgBox "Plug 'n Play Devices has been Enabled", vbInformation, "Success"
cmd8:
End Sub

Private Sub Command9_Click()
On Error GoTo cmd9
If Me.MousePointer = 14 Then
MsgBox Command9.ToolTipText, vbInformation, "Help!"
strtextprt = Command9.ToolTipText
Me.MousePointer = 0
Exit Sub
End If
strtextprt = Command9.ToolTipText
strText = String(30, " ") + strtextprt
If u2.value = 1 Then _
b = SaveValue("HKEY_CURRENT_USER", "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoTrayItemsDisplay", 1, "Dword")
If u1.value = 1 Then _
b = SaveValue("HKEY_LOCAL_MACHINE", "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoTrayItemsDisplay", 1, "Dword")
MsgBox "Tray Icons has been Disabled", vbInformation, "Success"
cmd9:
End Sub

Private Sub Form_Load()
strtextprt = " It's a utility to disable MSWindows features for security."
strText = String(30, " ") + strtextprt
End Sub

Private Sub statuslbl_Click()
If Me.MousePointer = 14 Then
MsgBox "It shows status report of current activity.", vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If

End Sub

Private Sub Timer1_Timer()
strText = Mid(strText, 2) & Left(strText, 1)
    statuslbl.Caption = "Status : " & strText
End Sub

