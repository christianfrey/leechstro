Attribute VB_Name = "ModRegedit"

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
        Public Const REG_SZ = 1 ' Unicode nul terminated String
        Public Const REG_DWORD = 4 ' 32-bit number


Public Function getstring(Hkey As Long, strPath As String, strValue As String)
        'EXAMPLE:
        '
        'text1.text = getstring(HKEY_CURRENT_USE
        '         R, "Software\VBW\Registry", "String")
        '
        Dim keyhand As Long
        Dim datatype As Long
        Dim lResult As Long
        Dim strBuf As String
        Dim lDataBufSize As Long
        Dim intZeroPos As Integer
        r = RegOpenKey(Hkey, strPath, keyhand)
        lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)


        If lValueType = REG_SZ Then
                strBuf = String(lDataBufSize, " ")
                lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)


                If lResult = ERROR_SUCCESS Then
                        intZeroPos = InStr(strBuf, Chr$(0))


                        If intZeroPos > 0 Then
                                getstring = Left$(strBuf, intZeroPos - 1)
                        Else
                                getstring = strBuf
                        End If
                End If
        End If
End Function

Public Sub savestring(Hkey As Long, strPath As String, strValue As String, strData As String)
        'EXAMPLE:
        '
        'Call savestring(HKEY_CURRENT_USER, "Sof
        '         tware\VBW\Registry", "String", text1.tex
        '         t)
        '
        Dim keyhand As Long
        Dim r As Long
        r = RegCreateKey(Hkey, strPath, keyhand)
        r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strData, Len(strData))
        r = RegCloseKey(keyhand)
End Sub

Public Function DeleteValue(ByVal Hkey As Long, ByVal strPath As String, ByVal strValue As String)
        'EXAMPLE:
        '
        'Call DeleteValue(HKEY_CURRENT_USER, "So
        '         ftware\VBW\Registry", "Dword")
        '
        Dim keyhand As Long
        r = RegOpenKey(Hkey, strPath, keyhand)
        r = RegDeleteValue(keyhand, strValue)
        r = RegCloseKey(keyhand)
End Function


