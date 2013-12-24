VERSION 5.00
Begin VB.Form sysetn02a 
   BorderStyle     =   0  'None
   Caption         =   "Kool System Settings"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7110
   Icon            =   "sysetn02a.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4500
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C000&
      Height          =   975
      Left            =   120
      TabIndex        =   19
      Top             =   2040
      Width           =   3135
      Begin VB.CommandButton Command8 
         Caption         =   "BackImage for Explorer Toolbar"
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
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "Wanna change backgorund Image for Explorer Toolbar."
         Top             =   120
         Width           =   2895
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Restore Explorer Toolbar"
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
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Wanna restore backgorund Image for Explorer Toolbar."
         Top             =   600
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      Height          =   975
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   3135
      Begin VB.CommandButton Command10 
         Caption         =   "Restore LogonScreen"
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
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   2895
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Change LogonScreen"
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
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Wanna change the login screen of WindowsXP."
         Top             =   120
         Width           =   2895
      End
   End
   Begin VB.TextBox fpath 
      Height          =   285
      Left            =   1080
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   4080
   End
   Begin VB.CommandButton Command11 
      Caption         =   "?"
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
      Left            =   6720
      TabIndex        =   14
      Top             =   4200
      Width           =   375
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Caption         =   "Kool Desky"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   4095
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   3375
      Begin VB.CommandButton Command6 
         Caption         =   "Rename 'My Network Places'"
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
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Wanna change default name of My Network Places."
         Top             =   1680
         Width           =   3135
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Rename 'Recycle Bin'"
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
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Wanna change default name of Recycle Bin."
         Top             =   1200
         Width           =   3135
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Rename 'My Documents'"
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
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Wanna change default name of My Documents folder."
         Top             =   720
         Width           =   3135
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Rename 'My Computer'"
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
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Wanna change default name of My Computer."
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   4095
      Left            =   3360
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.CommandButton Command2 
         Caption         =   "Remove Selected"
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
         Left            =   1920
         TabIndex        =   9
         Top             =   3480
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Apply Selected"
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
         Left            =   120
         TabIndex        =   8
         Top             =   3480
         Width           =   1695
      End
      Begin VB.CheckBox s5 
         BackColor       =   &H00FF8080&
         Caption         =   "Optimise Hard Disk when PC is Idle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Optimise Hard Disk when PC is Idle"
         Top             =   2880
         Width           =   3495
      End
      Begin VB.CheckBox s4 
         BackColor       =   &H00FF8080&
         Caption         =   "Hide Start Menu's Shutdown Option"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Hide Start Menu's Shutdown Option"
         Top             =   2280
         Width           =   3375
      End
      Begin VB.CheckBox s3 
         BackColor       =   &H00FF8080&
         Caption         =   "Disable Windows CD Burning"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Disable Windows CD Burning"
         Top             =   1680
         Width           =   3135
      End
      Begin VB.CheckBox s2 
         BackColor       =   &H00FF8080&
         Caption         =   "Increase the USB Device Polling Interval"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   495
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Increase the USB Device Polling Interval"
         Top             =   960
         Width           =   3255
      End
      Begin VB.CheckBox s1 
         BackColor       =   &H00FF8080&
         Caption         =   "Start 'Command Prompt' from directories by right click on Folder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   495
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Start 'Command Prompt' from directories by right click on Folder"
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Label statuslbl 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   4200
      Width           =   6735
   End
End
Attribute VB_Name = "sysetn02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strText, strtextprt As String

Private Sub Command1_Click()
On Error GoTo cmd1
Dim b As Boolean
If s1.value = 1 Then
b = CreateKey("HKEY_CLASSES_ROOT", "Folder\shell\&Open_DOS_Prompt_Here")
b = SaveValue("HKEY_CLASSES_ROOT", "Folder\shell\&Open_DOS_Prompt_Here", "", "&Open DOS Prompt here", "string")
b = CreateKey("HKEY_CLASSES_ROOT", "Folder\shell\&Open_DOS_Prompt_Here\command")
b = SaveValue("HKEY_CLASSES_ROOT", "Folder\shell\&Open_DOS_Prompt_Here\command", "", "cmd.exe /k cd %1", "string")
End If
If s2.value = 1 Then
b = SaveValue("HKEY_LOCAL_MACHINE", "SYSTEM\CurrentControlSet\Control\Class\{36FC9E60-C465-11CF-8056-444553540000}\0000", "IdleEnable", "1", "dword")
End If
If s3.value = 1 Then
b = SaveValue("HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer", "NoCDBurning", "1", "dword")
End If
If s4.value = 1 Then
b = SaveValue("HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer", "NoClose", "1", "dword")
End If
If s5.value = 1 Then
b = SaveValue("HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Windows\CurrentVersion\OptimalLayout", "EnableAutoLayout", "1", "dword")
End If
MsgBox "Selected FX have been applied effectively.", vbInformation, "Success"
cmd1:
End Sub

Private Sub Command10_Click()
On Error GoTo cmd10
If Me.MousePointer = 14 Then
MsgBox "Wanna change the login screen of WindowsXP.", vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
fpath.Text = "logonui.exe"
If Mid(fpath.Text, Len(fpath.Text) - 3, 4) = ".exe" Then
b = SaveValue("HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "UIHost", fpath.Text, "string")
If b = True Then
MsgBox "Logon Restored :: LogIn again for proper effect.", vbInformation, "Restore LogonScreen"
strtextprt = Command10.ToolTipText
strText = String(30, " ") + strtextprt
End If
End If
cmd10:
End Sub

Private Sub Command11_Click()
Me.MousePointer = 14
End Sub

Private Sub Command2_Click()
On Error GoTo cmd2
Dim b As Boolean
If s1.value = 1 Then
 b = DeleteKey("HKEY_CLASSES_ROOT", "Folder\shell\&Open_DOS_Prompt_Here\command")
 b = DeleteKey("HKEY_CLASSES_ROOT", "Folder\shell\&Open_DOS_Prompt_Here")
End If
If s2.value = 1 Then
b = DeleteValue("HKEY_LOCAL_MACHINE", "SYSTEM\CurrentControlSet\Control\Class\{36FC9E60-C465-11CF-8056-444553540000}\0000", "IdleEnable")
End If
If s3.value = 1 Then
b = SaveValue("HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer", "NoCDBurning", "0", "dword")
End If
If s4.value = 1 Then
b = SaveValue("HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer", "NoClose", "0", "dword")
End If
If s5.value = 1 Then
b = SaveValue("HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Windows\CurrentVersion\OptimalLayout", "EnableAutoLayout", "0", "dword")
End If
MsgBox "Selected FX have been removed effectively.", vbInformation, "Success"
cmd2:
End Sub

Private Sub Command3_Click()
On Error GoTo cmd3
If Me.MousePointer = 14 Then
MsgBox "Wanna change default name of My Computer.", vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
naam = "My Computer"
naam = InputBox("Enter new name for My Computer", "My Computer's New Name", "My Computer")
If naam <> "" Then
b = SaveValue("HKEY_CURRENT_USER", "Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}", "", naam, "string")
If b = True Then
MsgBox "Name Changed :: LogIn again for proper effect."
strtextprt = Command3.ToolTipText
strText = String(30, " ") + strtextprt
End If
End If
cmd3:
End Sub

Private Sub Command4_Click()
On Error GoTo cmd4
If Me.MousePointer = 14 Then
MsgBox "Wanna change default name of My Documents folder.", vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
naam = "My Computer"
naam = InputBox("Enter new name for My Documents", "My Documents' New Name", "My Documents")
If naam <> "" Then
b = SaveValue("HKEY_CURRENT_USER", "Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{450D8FBA-AD25-11D0-98A8-0800361B1103}", "", naam, "string")
If b = True Then
MsgBox "Name Changed :: LogIn again for proper effect."
strtextprt = Command4.ToolTipText
strText = String(30, " ") + strtextprt
End If
End If
cmd4:
End Sub

Private Sub Command5_Click()
On Error GoTo cmd5
If Me.MousePointer = 14 Then
MsgBox "Wanna change default name of Recycle Bin.", vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
naam = "Recycle Bin"
naam = InputBox("Enter new name for Recycle Bin", "Recycle Bin's New Name", "Recycle Bin")
If naam <> "" Then
b = SaveValue("HKEY_CURRENT_USER", "Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", "", naam, "string")
If b = True Then
MsgBox "Name Changed :: LogIn again for proper effect."
strtextprt = Command5.ToolTipText
strText = String(30, " ") + strtextprt
End If
End If
cmd5:
End Sub

Private Sub Command6_Click()
On Error GoTo cmd6
If Me.MousePointer = 14 Then
MsgBox "Wanna change default name of My Network Places.", vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
naam = "My Network Places"
naam = InputBox("Enter new name for My Network Places", "My Network Places' New Name", "My Network Places")
If naam <> "" Then
b = SaveValue("HKEY_CURRENT_USER", "Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{208D2C60-3AEA-1069-A2D7-08002B30309D}", "", naam, "string")
If b = True Then
MsgBox "Name Changed :: LogIn again for proper effect."
strtextprt = Command6.ToolTipText
strText = String(30, " ") + strtextprt
End If
End If
cmd6:
End Sub

Private Sub Command7_Click()
On Error GoTo cmd7
If Me.MousePointer = 14 Then
MsgBox "Wanna restore backgorund Image for Explorer Toolbar.", vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
b = SaveValue("HKEY_CURRENT_USER", "Software\Microsoft\Internet Explorer\Toolbar", "BackBitmapShell", "", "string")
If b = True Then
MsgBox "Toolbar Restored :: LogIn again for proper effect.", vbInformation, "Restore Toolbar"
strtextprt = Command7.ToolTipText
strText = String(30, " ") + strtextprt
End If
cmd7:
End Sub

Private Sub Command8_Click()
On Error GoTo cmd8
If Me.MousePointer = 14 Then
MsgBox "Wanna change backgorund Image for Explorer Toolbar.", vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
MsgBox "Please select a bitmap file only.", vbInformation, "Required"
fileb.Drive1.Drive = Mid(App.Path, 1, 3)
If Len(App.Path) <> 3 Then
fileb.Dir1.Path = App.Path & "\res"
Else
fileb.Dir1.Path = App.Path & "res"
End If
fileb.Show (vbModal)
If Mid(fpath.Text, Len(fpath.Text) - 3, 4) = ".bmp" Then
Dim FileSystemObject As Object
Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
FileSystemObject.CopyFile fpath.Text, SystemDirectory & "\exp_toolbar.bmp"
b = SaveValue("HKEY_CURRENT_USER", "Software\Microsoft\Internet Explorer\Toolbar", "BackBitmapShell", SystemDirectory & "\exp_toolbar.bmp", "string")
If b = True Then
MsgBox "Image Changed :: LogIn again for proper effect.", vbInformation, "New Toolbar Image"
strtextprt = Command8.ToolTipText
strText = String(30, " ") + strtextprt
End If
Else
MsgBox "File selected was not of Bitmap Format", vbCritical, "Not Changed"
End If
cmd8:
End Sub

Private Sub Command9_Click()
On Error GoTo cmd9
If Me.MousePointer = 14 Then
MsgBox "Wanna change the login screen of WindowsXP.", vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
MsgBox "Please select a valid logon file only. Otherwise system would be forced to classic logon.", vbInformation, "Required"
fileb.Drive1.Drive = Mid(App.Path, 1, 3)
If Len(App.Path) <> 3 Then
fileb.Dir1.Path = App.Path & "\res"
Else
fileb.Dir1.Path = App.Path & "res"
End If
fileb.Show (vbModal)
If Mid(fpath.Text, Len(fpath.Text) - 3, 4) = ".exe" Then
Dim FileSystemObject As Object
Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
FileSystemObject.CopyFile fpath.Text, SystemDirectory & "\AlogonBuiK.exe"
b = SaveValue("HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "UIHost", SystemDirectory & "\AlogonBuiK.exe", "string")
If b = True Then
MsgBox "Logon Screen Changed :: LogIn again for proper effect.", vbInformation, "New LogonScreen"
strtextprt = Command9.ToolTipText
strText = String(30, " ") + strtextprt
End If
End If
cmd9:
End Sub

Private Sub Form_Load()
strtextprt = " It's a utility to give kool effects to ur system."
strText = String(90, " ") + strtextprt
End Sub

Private Sub s1_Click()
If Me.MousePointer = 14 Then
MsgBox "This would add a new Menu Item to Right Click of all" _
& vbCrLf & "Folders enabling user to start Command Prompt at that Folder's Location.", _
vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
End Sub

Private Sub s2_Click()
If Me.MousePointer = 14 Then
MsgBox "This would increase the rate of USB Device Polling which." _
& vbCrLf & "would let Windows identify the attached USB device sooner.", vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
End Sub

Private Sub s3_Click()
If Me.MousePointer = 14 Then
MsgBox "This would disable the Windows Inbuilt CD Burning Feature" _
& vbCrLf & "to control the your data transfer more better way.", vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
End Sub

Private Sub s4_Click()
If Me.MousePointer = 14 Then
MsgBox "This would hide the Shutdown Option from Start Menu" _
& vbCrLf & "if you want to do it for some reason.", vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
End Sub

Private Sub s5_Click()
If Me.MousePointer = 14 Then
MsgBox "Enables your OS to optimise its HDD in free time itself" _
& vbCrLf & "to keep your HDD optimised state.", vbInformation, "Help!"
Me.MousePointer = 0
Exit Sub
End If
End Sub

Private Sub Timer1_Timer()
strText = Mid(strText, 2) & Left(strText, 1)
    statuslbl.Caption = "Status : " & strText
End Sub
