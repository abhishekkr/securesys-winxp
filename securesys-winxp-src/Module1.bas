Attribute VB_Name = "Module1"
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Dim i As Integer, f1(5) As Form


Public Sub load()
On Error GoTo NoLoadErr

Set f1(0) = New Form2
f1(0).Top = 500
f1(0).Left = Form1.Left '+ (i * 1860)
f1(0).Command1.Caption = "Not Set"
f1(0).Show
'Me.Visible = False
For i = 1 To 5
Set f1(i) = New Form2
f1(i).Top = f1(0).Top + (i * 1500)
f1(i).Left = f1(0).Left  '+ (i * 1860)
f1(i).Command1.Caption = "Not Set"
f1(i).Show
Next i
f1(0).Command1.Caption = "Windows StartUp Controller"
f1(1).Command1.Caption = "Kool System Settings"
f1(2).Command1.Caption = "Internet Explorer Settings"
f1(3).Command1.Caption = "Security Settings"
f1(4).Command1.Caption = "Settings 'n Help"
f1(5).Command1.Caption = "Exit"
Exit Sub
NoLoadErr:
MsgBox "There was some error in loading the Application." & vbCrLf _
& "Some components needed might be missing on your System.", _
vbOKOnly, "Problem :"
End Sub

Public Sub StayOnTop()
If f1(0).Check1.value = 0 Then
  f1(0).Check1.value = 1
  SetWindowPos f1(0).hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
  SetWindowPos f1(1).hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
  SetWindowPos f1(2).hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
  SetWindowPos f1(3).hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
  SetWindowPos f1(4).hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
  SetWindowPos f1(5).hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
ElseIf f1(0).Check1.value = 1 Then
  f1(0).Check1.value = 0
  SetWindowPos f1(0).hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
  SetWindowPos f1(1).hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
  SetWindowPos f1(2).hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
  SetWindowPos f1(3).hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
  SetWindowPos f1(4).hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
  SetWindowPos f1(5).hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End If
End Sub

Public Sub unloadall()
For i = 1 To 5
Unload f1(i - 1)
Next i
Unload Form1
Unload frmset
Unload loader
End
End Sub

Public Sub movetabs()
f1(0).Top = Form1.Top + (1000)
f1(0).Left = Form1.Left '+ (i * 1860)
f1(0).Show
'Me.Visible = False
For i = 1 To 5
f1(i).Top = f1(0).Top + (i * 900)
f1(i).Left = f1(0).Left  '+ (i * 1860)
f1(i).Show
Next i
End Sub

