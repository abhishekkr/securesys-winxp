Attribute VB_Name = "Module3"
Option Explicit

Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hkey As Long, ByVal lpValueName As String) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hkey As Long, ByVal lpSubKey As String) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbdata As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbdata As Long) As Long
Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbdata As Long) As Long
Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long

'Security attribute type
Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

'Handles of key(Hives)
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006

'Security access constant
Const KEY_QUERY_VALUE = &H1&
Const KEY_SET_VALUE = &H2&
Const KEY_CREATE_SUB_KEY = &H4&
Const KEY_ENUMERATE_SUB_KEYS = &H8&
Const KEY_NOTIFY = &H10&
Const KEY_CREATE_LINK = &H20&
Const READ_CONTROL = &H20000
Const WRITE_DAC = &H40000
Const WRITE_OWNER = &H80000
Const SYNCHRONIZE = &H100000
Const STANDARD_RIGHTS_ALL = &H1F0000
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const STANDARD_RIGHTS_READ = READ_CONTROL
Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Const KEY_EXECUTE = KEY_READ
Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

'Error Constant
Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1009&
Const ERROR_BADKEY = 1010&
Const ERROR_CANTOPEN = 1011&
Const ERROR_CANTREAD = 1012&
Const ERROR_CANTWRITE = 1013&
Const ERROR_OUTOFMEMORY = 14&
Const ERROR_INVALID_PARAMETER = 87&
Const ERROR_ACCESS_DENIED = 5&
Const ERROR_NO_MORE_ITEMS = 259&
Const ERROR_MORE_DATA = 234&

'Other Constant
Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted
Const REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number
Const REG_BINARY = 3                     ' Free form binary

'window Up time API
Public Declare Function GetTickCount& Lib "kernel32" ()
'===================================================
'Function: SystemDirectory

Public Declare Function GetSystemDirectory& Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long)

Public Function SystemDirectory()
  Dim strBuffer As String * 260

  GetSystemDirectory strBuffer, 260

  SystemDirectory = Left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
End Function

'***********************************************************
'###########################################################
'***********************************************************
'This function convert string value to long
Public Function hkeyf(ByVal hive As String) As Long
Dim hkey As Long
If hive = "HKEY_CLASSES_ROOT" Then
hkey = HKEY_CLASSES_ROOT
ElseIf hive = "HKEY_CURRENT_USER" Then
hkey = HKEY_CURRENT_USER
ElseIf hive = "HKEY_LOCAL_MACHINE" Then
hkey = HKEY_LOCAL_MACHINE
ElseIf hive = "HKEY_USERS" Then
hkey = HKEY_USERS
ElseIf hive = "HKEY_PERFORMANCE_DATA" Then
hkey = HKEY_PERFORMANCE_DATA
ElseIf hive = "HKEY_CURRENT_CONFIG" Then
hkey = HKEY_CURRENT_CONFIG
ElseIf hive = "HKEY_DYN_DATA" Then
hkey = HKEY_DYN_DATA
End If
hkeyf = hkey
End Function

'This function return true if sucess otherwise false
Public Function Message(i As Long) As Boolean
If i = ERROR_SUCCESS Then
Message = True
Else
Message = False
End If
End Function

'**********************************************************************
'Delete a Registry value
'**********************************************************************
Public Function DeleteValue(ByVal hive As String, ByVal key As String, ByVal valuename As String) As Boolean
Dim hkey As Long
Dim i As Long
Dim h As Long
hkey = hkeyf(hive)
RegOpenKeyEx hkey, key, 0&, KEY_ALL_ACCESS, h
i = RegDeleteValue(h, valuename)
DeleteValue = Message(i)
RegCloseKey h
End Function

'***********************************************************
' It create or open existing key
'***********************************************************
Public Function CreateKey(ByVal hive As String, ByVal key As String) As Boolean
Dim hkey As Long
Dim i As Long
Dim j As Long
Dim hkeynew As Long
hkey = hkeyf(hive)
' The next three lines give default values for secattr
Dim s As SECURITY_ATTRIBUTES
s.lpSecurityDescriptor = 0  ' default security level
s.bInheritHandle = True     ' might as well allow it
s.nLength = Len(s)          ' store size of variable
i = RegCreateKeyEx(hkey, key, 0&, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, s, hkeynew, j)
' if created, j = 1; if key exists & just open , j = 2
CreateKey = Message(i)
RegCloseKey hkeynew
End Function


'**********************************************************************
'It show all sub key of a registry key
'**********************************************************************
Public Function listKey(ByVal hive As String, ByVal key As String)
Dim hkey As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim outname As String
Dim outsize As Long
Dim outclass As String
Dim outclasssize As Long
Dim index As Long
Dim ss As String
Dim LTime As FILETIME
hkey = hkeyf(hive)
i = RegOpenKeyEx(hkey, key, 0&, KEY_ALL_ACCESS, k)
If i = ERROR_SUCCESS Then
Do
outname = Space(1024)  'This is the buffer value
outsize = 1024         'This is the buffer data
outclass = Space(1024)  'This is the buffer value
outclasssize = 1024         'This is the buffer data
j = RegEnumKeyEx(k, index, outname, outsize, 0&, outclass, outclasssize, LTime)
If j = ERROR_SUCCESS Then
    index = index + 1
    ss = ""
    ss = key & "\" & Left(outname, outsize)
    ss = Trim(ss)
    MsgBox ss
End If
Loop Until j <> ERROR_SUCCESS
End If
RegCloseKey k
End Function

'***********************************************************
'It return the datavalue of the key
'***********************************************************
Public Function GetValue(ByVal hive As String, ByVal key As String, ByVal valuename As String) As String
Dim hkey As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim h As Long
Dim ss As String
Dim ii As Integer
Dim outtype As Long
Dim outdata As String
Dim outsize As Long
hkey = hkeyf(hive)     'Return the hkey value
outdata = Space(1024)  'This is the buffer value
outsize = 1024         'This is the buffer data
i = RegOpenKeyEx(hkey, key, 0, KEY_ALL_ACCESS, h)
If i = ERROR_SUCCESS Then
j = RegQueryValueEx(h, valuename, 0&, outtype, outdata, outsize)
outdata = Trim(outdata)
Select Case outtype
Case REG_SZ
    GetValue = outdata
Case REG_DWORD
     Dim dworddata As Long
     j = RegQueryValueExLong(h, valuename, 0&, outtype, dworddata, outsize)
     If j = ERROR_SUCCESS Then
          GetValue = dworddata
     Else
          GetValue = ""
     End If
Case REG_BINARY
     ss = ""
     For ii = 1 To Len(outdata)
        If Len(Hex(Asc(Mid(outdata, ii, 1)))) = 1 Then
        ss = ss & "0" & Hex(Asc(Mid(outdata, ii, 1)))
        Else
        ss = ss & Hex(Asc(Mid(outdata, ii, 1)))
        End If
     Next
     GetValue = ss
End Select
Else
GetValue = ""
End If
RegCloseKey h
End Function

'***********************************************************
'It Save data value of key
'***********************************************************
Public Function SaveValue(ByVal hive As String, ByVal key As String, ByVal valuename As String, ByVal value As Variant, ByVal datatype As String) As Boolean
Dim hkey As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim h As Long
Dim ss As String
Dim s As String
Dim ii As Long
Dim f As Long
hkey = hkeyf(hive)     'Return the hkey value
'i = RegOpenKeyEx(hKey, key, 0, KEY_ALL_ACCESS, h)
' The next three lines give default values for secattr
    Dim a As SECURITY_ATTRIBUTES
    a.lpSecurityDescriptor = 0  ' default security level
    a.bInheritHandle = True     ' might as well allow it
    a.nLength = Len(a)          ' store size of variable
i = RegCreateKeyEx(hkey, key, 0&, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, a, h, k)
If i = ERROR_SUCCESS Then
value = Trim(value)
Select Case LCase(datatype)
Case "string"     'For string values
    j = RegSetValueEx(h, valuename, 0&, REG_SZ, value, Len(value))
Case "binary"     'For binary value
    For ii = 1 To Len(value) Step 2
         ss = ""
         ss = Mid(value, ii, 2)
         ss = "&H" & ss
         s = s & Chr(Val(CDec(ss)))
    Next
    value = s
    j = RegSetValueExString(h, valuename, 0&, REG_BINARY, value, Len(value))
Case "dword"     'For dword value
    j = RegSetValueExLong(h, valuename, 0&, REG_DWORD, value, 4)
End Select
SaveValue = Message(j)
Else
SaveValue = False
End If
RegCloseKey h
End Function

'**********************************************************************
'It delete a key.
'**********************************************************************

Public Function DeleteKey(ByVal hive As String, ByVal key As String) As Boolean
Dim hkey As Long
Dim i As Long
Dim k As Long
hkey = hkeyf(hive)
i = RegOpenKeyEx(hkey, key, 0&, KEY_ALL_ACCESS, k)
If i = ERROR_SUCCESS Then
i = RegDeleteKey(hkey, key)
DeleteKey = Message(i)
End If
RegCloseKey k
End Function


'**********************************************************************
'It return all regestry valuename (data, data type, value name)
'**********************************************************************
Public Function listkeyvalue(ByVal hive As String, ByVal key As String)
Dim h As Long
Dim hkey As Long
Dim cenum As Long
Dim i As Long
Dim strnamebuff As String
Dim sizebuff As Long
Dim lngtype As Long
Dim abytdata(1 To 2048) As Byte
Dim j As Long
Dim valuename As String
Dim value As String
hkey = hkeyf(hive)
i = RegOpenKeyEx(hkey, key, 0&, KEY_ALL_ACCESS, h)
If i = ERROR_SUCCESS Then
Do
strnamebuff = Space$(1024)
sizebuff = Len(strnamebuff)
Erase abytdata
j = UBound(abytdata)
i = RegEnumValue(h, cenum, strnamebuff, sizebuff, 0&, lngtype, abytdata(1), j)
'**********Note:****************
'lngtype--------- return the data type of the value
'strnamebuff----- return name of the value
'value----------- return data value
If i = ERROR_SUCCESS Then
valuename = Left(strnamebuff, sizebuff)
value = ""
value = GetValue(hive, key, valuename)
value = Trim(value)
    Select Case lngtype
    Case REG_BINARY
    MsgBox "Value name=  " & valuename & "; Data type=Binary; Value= " & value
    Case REG_DWORD
    MsgBox "Value name= " & valuename & "; Data type=Dword; Value= " & value
    Case REG_SZ
    MsgBox "Value name= " & valuename & "Data type=String; Value= " & value
    End Select

End If
cenum = cenum + 1
Loop Until i <> 0
i = RegCloseKey(h)
End If
End Function

'************************************* End **********************************************

'--------now work starts-----------------'
Public Function listRunkeyvalue(ByVal hive As String, ByVal key As String)
Dim h As Long
Dim hkey As Long
Dim cenum As Long
Dim i As Long
Dim strnamebuff As String
Dim sizebuff As Long
Dim lngtype As Long
Dim abytdata(1 To 2048) As Byte
Dim j As Long
Dim valuename As String
Dim value As String
hkey = hkeyf(hive)
i = RegOpenKeyEx(hkey, key, 0&, KEY_ALL_ACCESS, h)
If i = ERROR_SUCCESS Then
Do
strnamebuff = Space$(1024)
sizebuff = Len(strnamebuff)
Erase abytdata
j = UBound(abytdata)
i = RegEnumValue(h, cenum, strnamebuff, sizebuff, 0&, lngtype, abytdata(1), j)
'**********Note:****************
'lngtype--------- return the data type of the value
'strnamebuff----- return name of the value
'value----------- return data value
If i = ERROR_SUCCESS Then
valuename = Left(strnamebuff, sizebuff)
value = ""
value = GetValue(hive, key, valuename)
value = Trim(value)
    Select Case lngtype
    Case REG_BINARY
    'MsgBox "Value name=  " & valuename & "; Data type=Binary; Value= " & value
    Case REG_DWORD
    'MsgBox "Value name= " & valuename & "; Data type=Dword; Value= " & value
    Case REG_SZ
    'MsgBox "Value name= " & valuename & "Data type=String; Value= " & value
    End Select
    
winstart01a.HKLMRunListKey.AddItem (valuename)
End If
cenum = cenum + 1
Loop Until i <> 0
i = RegCloseKey(h)
End If
End Function

Public Function listIEbarvalue(ByVal hive As String, ByVal key As String)
Dim h As Long
Dim hkey As Long
Dim cenum As Long
Dim i As Long
Dim strnamebuff As String
Dim sizebuff As Long
Dim lngtype As Long
Dim abytdata(1 To 2048) As Byte
Dim j As Long
Dim valuename As String
Dim value As String
hkey = hkeyf(hive)
i = RegOpenKeyEx(hkey, key, 0&, KEY_ALL_ACCESS, h)
If i = ERROR_SUCCESS Then
Do
strnamebuff = Space$(1024)
sizebuff = Len(strnamebuff)
Erase abytdata
j = UBound(abytdata)
i = RegEnumValue(h, cenum, strnamebuff, sizebuff, 0&, lngtype, abytdata(1), j)
'**********Note:****************
'lngtype--------- return the data type of the value
'strnamebuff----- return name of the value
'value----------- return data value
If i = ERROR_SUCCESS Then
valuename = Left(strnamebuff, sizebuff)
value = ""
value = GetValue(hive, key, valuename)
value = Trim(value)
    Select Case lngtype
    Case REG_BINARY
    'MsgBox "Value name=  " & valuename & "; Data type=Binary; Value= " & value
    Case REG_DWORD
    'MsgBox "Value name= " & valuename & "; Data type=Dword; Value= " & value
    Case REG_SZ
    'MsgBox "Value name= " & valuename & "Data type=String; Value= " & value
    End Select
    
ieset03a.ListIEbar.AddItem (value)
ieset03a.listtmp.AddItem (valuename)
End If
cenum = cenum + 1
Loop Until i <> 0
i = RegCloseKey(h)
End If
End Function

Public Function listIEMENUKey(ByVal hive As String, ByVal key As String)
Dim hkey As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim outname As String
Dim outsize As Long
Dim outclass As String
Dim outclasssize As Long
Dim index As Long
Dim ss As String
Dim LTime As FILETIME
hkey = hkeyf(hive)
i = RegOpenKeyEx(hkey, key, 0&, KEY_ALL_ACCESS, k)
If i = ERROR_SUCCESS Then
Do
outname = Space(1024)  'This is the buffer value
outsize = 1024         'This is the buffer data
outclass = Space(1024)  'This is the buffer value
outclasssize = 1024         'This is the buffer data
j = RegEnumKeyEx(k, index, outname, outsize, 0&, outclass, outclasssize, LTime)
If j = ERROR_SUCCESS Then
    index = index + 1
    ss = ""
    ss = key & "\" & Left(outname, outsize)
    ss = Trim(ss)
    ieset03a.ListIEbar.AddItem (outname)
    'ieset03a.listtmp.AddItem (key)
End If
Loop Until j <> ERROR_SUCCESS
End If
RegCloseKey k
End Function

'===================================================
'Function: GetWindowsUpTime

Public Function GWinUp()
  GWinUp = GetTickCount / 1000
End Function

