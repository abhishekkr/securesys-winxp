Attribute VB_Name = "Module2"
Option Explicit

Public Const LWA_BOTH = 3
Public Const GWL_EXSTYLE = (-20)
Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2
Public Const ULW_COLORKEY = &H1
Public Const ULW_ALPHA = &H2
Public Const ULW_OPAQUE = &H4
Public Const WS_EX_LAYERED = &H80000

' To corner round (by implemented)
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long

' To Transparency
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function UpdateLayeredWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


Public Function MakeTransparent(ByVal hwnd As Long, Perc As Integer) As Long
  Dim lngMsg As Long

  On Error Resume Next
  If Perc < 0 Or Perc > 190 Then
    MakeTransparent = 1
  Else
    lngMsg = GetWindowLong(hwnd, GWL_EXSTYLE)
    lngMsg = lngMsg Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, lngMsg
    SetLayeredWindowAttributes hwnd, 0, Perc, LWA_ALPHA
    MakeTransparent = 0
  End If
  
  If Err Then MakeTransparent = 2
End Function
