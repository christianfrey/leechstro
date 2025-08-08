Attribute VB_Name = "ModIniFile"
'
' Module permettant de modifier le fichier "configurations.ini" avec les
' paramètres des options de "frmOptions".
'

Option Explicit

Private Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName$) As Long

Public Function GetIni(Section As String, Variable As String, Fichier As String) As String
Dim strRetour As String
strRetour = String(255, Chr(0))
Dim Longueur As Integer
Longueur = GetPrivateProfileString(Section, Variable, "", strRetour, Len(strRetour), Fichier)
GetIni = Left$(strRetour, Longueur)
End Function

Function WriteIni(Section As String, Variable As String, Valeur As String, Fichier As String, Optional nopefface As Integer) As Integer
If nopefface = 0 Then
WriteIni Section, Variable, "", Fichier, 1
End If
WritePrivateProfileString Section, Variable, Valeur, Fichier
End Function
